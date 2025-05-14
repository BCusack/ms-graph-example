import asyncio
import os
import base64  # Added for sharing token
import aiofiles  # Added for async file operations
from azure.identity.aio import ClientSecretCredential
from msgraph.graph_service_client import GraphServiceClient
from kiota_abstractions.api_error import APIError  # General Kiota API error
from msgraph.generated.service_principals.service_principals_request_builder import (
    ServicePrincipalsRequestBuilder,
)
import aiohttp  # For making the actual upload PUT request

# Set your Azure AD app details (or use environment variables)
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")


# Helper function to create the sharing token for the Graph API
def get_sharing_token(share_url: str) -> str:
    """
    Encodes a SharePoint URL into the format required by the Graph API for shared items.
    u! + base64urlencode(share_url)
    """
    # Base64 encode the URL
    encoded_url_bytes = base64.urlsafe_b64encode(share_url.encode("utf-8"))
    # Convert bytes to string and remove padding
    encoded_url_str = encoded_url_bytes.decode("utf-8").rstrip("=")
    return f"u!{encoded_url_str}"


async def download_sharepoint_url_as_pdf(
    client: GraphServiceClient, share_url: str, output_file_path: str
):
    """
    Downloads a document from a SharePoint sharing URL and saves it as a PDF.
    """
    if not share_url:
        print("Error: SharePoint URL is empty.")
        return
    if not output_file_path:
        print("Error: Output file path is empty.")
        return

    try:
        sharing_token = get_sharing_token(share_url)
        print(
            f"Generated sharing token: {sharing_token[:50]}..."
        )  # Print part of token for brevity

        # 1. Resolve the sharing URL to a DriveItem
        print("Resolving SharePoint URL to DriveItem...")  # Corrected f-string
        shared_item = await client.shares.by_shared_drive_item_id(
            sharing_token
        ).drive_item.get()

        if (
            not shared_item
            or not shared_item.id
            or not shared_item.parent_reference
            or not shared_item.parent_reference.drive_id
        ):
            print(
                "Could not resolve the sharing URL to a valid DriveItem. Ensure the link is correct and accessible."
            )  # Corrected f-string
            if shared_item:
                print(
                    f"Resolved item ID: {shared_item.id}, Drive ID: {shared_item.parent_reference.drive_id if shared_item.parent_reference else 'N/A'}"
                )
            return

        drive_id = shared_item.parent_reference.drive_id
        item_id = shared_item.id
        file_name = shared_item.name or "downloaded_file"
        print(
            f"Successfully resolved DriveItem: Name: '{file_name}', ID: '{item_id}', Drive ID: '{drive_id}'"
        )

        # 2. Request the DriveItem content as PDF
        print(f"Requesting content of '{file_name}' as PDF...")

        request_info = (
            client.drives.by_drive_id(drive_id)
            .items.by_drive_item_id(item_id)
            .content.to_get_request_information()
        )
        if request_info:
            request_info.url += "?format=pdf"
            pdf_content = await client.request_adapter.send_primitive_async(
                request_info, "bytes", None
            )
        else:
            print("Failed to create request information for content download.")
            return

        if pdf_content:
            print(
                f"Successfully downloaded content (size: {len(pdf_content)} bytes). Saving to '{output_file_path}'..."
            )
            # 3. Save the PDF content
            async with aiofiles.open(output_file_path, "wb") as f:
                await f.write(pdf_content)
            print(f"File saved successfully to {output_file_path}")
        else:
            print("Failed to download PDF content or content was empty.")

    except Exception as e:
        import traceback

        print(f"An error occurred during SharePoint file download: {e}")
        print(traceback.format_exc())


async def get_me(client: GraphServiceClient):
    """
    Fetches the current user's details.
    """
    try:
        me = await client.me.get()
        if me:
            print(f"User ID: {me.id}, Display Name: {me.display_name}")
        else:
            print("No user information found.")
    except Exception as e:
        print(f"An error occurred while fetching user details: {e}")


async def get_all_users(client: GraphServiceClient):
    """
    Fetches
    """
    try:
        users = await client.users.get()
        if users and users.value:
            for user in users.value:
                print(f"User ID: {user.id}, Display Name: {user.display_name}")
        else:
            print("No user information found.")
    except Exception as e:
        print(f"An error occurred while fetching user details: {e}")


async def main():
    # Ensure required environment variables are set
    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        print(
            "Error: TENANT_ID, CLIENT_ID, or CLIENT_SECRET is not set or is invalid. "
            "Please check your .env file and ensure they are correct full values."
        )
        return

    # Authenticate with client credentials
    credential = ClientSecretCredential(
        tenant_id=TENANT_ID,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET,
    )
    scopes = [
        "https://graph.microsoft.com/.default"
    ]  # Ensure Sites.Read.All is granted in Azure AD
    client = GraphServiceClient(credentials=credential, scopes=scopes)

    await get_all_users(client)

    try:
        print(
            f"Attempting to fetch service principal for application (Client ID: {CLIENT_ID})..."
        )
        # For app-only authentication, get the service principal associated with the application
        query_params = ServicePrincipalsRequestBuilder.ServicePrincipalsRequestBuilderGetQueryParameters(
            filter=f"appId eq '{CLIENT_ID}'"
        )
        request_configuration = ServicePrincipalsRequestBuilder.ServicePrincipalsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        service_principal_response = await client.service_principals.get(
            request_configuration=request_configuration
        )

        if service_principal_response and service_principal_response.value:
            app_service_principal = service_principal_response.value[0]
            print(
                f"Successfully authenticated. Application Display Name: {app_service_principal.display_name} (ID: {app_service_principal.id})"
            )
        elif service_principal_response:
            print(
                f"Could not find service principal for app ID: {CLIENT_ID}. Response: {service_principal_response}"
            )
        else:
            print(
                f"Failed to retrieve service principal for app ID: {CLIENT_ID}. No response."
            )

    except Exception as e:
        import traceback

        print(f"An error occurred: {e}")
        print(traceback.format_exc())


async def upload_file_to_sharepoint(
    client: GraphServiceClient,
    file_path: str,
    sharepoint_site_id: str,
    drive_name: str,
    folder_path: str = "",
):
    if not file_path or not os.path.isfile(file_path):
        print(f"Error: File '{file_path}' does not exist.")
        return

    # Get the drive (document library) ID
    try:
        drives = await client.sites.by_site_id(sharepoint_site_id).drives.get()
    except APIError as e:
        print(
            f"Graph API Error getting drives: Status {e.response_status_code}, {str(e)}"
        )
        return
    except Exception as e:
        print(f"Generic error getting drives: {e}")
        return

    drive_id = None
    drives_value = getattr(drives, "value", None)
    if drives and drives_value is not None:
        for drive in drives_value:
            if drive.name == drive_name:
                drive_id = drive.id
                break
    else:
        print(
            f"Could not retrieve drives for site '{sharepoint_site_id}'. Response: {drives}"
        )
        return
    if not drive_id:
        print(f"Drive '{drive_name}' not found.")
        return

    file_name = os.path.basename(file_path)

    clean_folder_path = folder_path.strip("/\\\\").replace("\\\\", "/")
    if clean_folder_path:
        upload_path_relative_to_root = f"{clean_folder_path}/{file_name}"
    else:
        upload_path_relative_to_root = file_name

    # Simplified request body
    request_body_dict = {"item": {"@microsoft.graph.conflictBehavior": "replace"}}

    # The item specifier for the API call, e.g., "root:/MyFolder/MyFile.txt"
    item_specifier = f"root:/{upload_path_relative_to_root.lstrip('/')}"

    print(
        f"Creating upload session for item specifier: {item_specifier} in drive {drive_id}"
    )

    try:
        upload_session_request_builder = (
            client.drives.by_drive_id(drive_id)
            .items.by_drive_item_id(item_specifier)
            .create_upload_session
        )
        upload_session = await upload_session_request_builder.post(
            request_body=request_body_dict
        )

        await (
            client.sites.by_site_id(site_id)
            .drive.root.item_with_path("Documents/MyFile.txt")
            .content.put(body=file_content)
        )

    except APIError as e:
        print(
            f"Graph API Error creating upload session: Status {e.response_status_code}"
        )
        print(f"Error details: {str(e)}")
        if hasattr(e, "error") and e.error and hasattr(e.error, "message"):
            print(f"Detailed message: {e.error.message}")
        return
    except Exception as e:
        print(f"Generic error creating upload session: {type(e).__name__} - {e}")
        return

    if not upload_session or not upload_session.upload_url:
        print("Failed to create upload session or upload URL not found.")
        return

    upload_url = upload_session.upload_url

    file_size = os.path.getsize(file_path)
    print(f"File size: {file_size} bytes. Uploading to: {upload_url}")

    async with aiohttp.ClientSession() as http_session:
        async with aiofiles.open(file_path, "rb") as f:
            content = await f.read()
            headers = {
                "Content-Length": str(file_size),
            }
            async with http_session.put(
                upload_url, data=content, headers=headers
            ) as response:
                if response.status >= 200 and response.status < 300:
                    print(f"File uploaded successfully. Status: {response.status}")
                    response_json = await response.json()
                    print(f"Server response: {response_json}")
                    return response_json
                else:
                    print(f"Upload failed. Status: {response.status}")
                    error_text = await response.text()
                    print(f"Error details: {error_text}")
                    return None


if __name__ == "__main__":
    asyncio.run(main())
