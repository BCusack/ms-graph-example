import asyncio
import os
import base64  # Added for sharing token
import aiofiles  # Added for async file operations

from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient

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


async def main():
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

    if not CLIENT_ID or not TENANT_ID or not CLIENT_SECRET:
        print(
            "Error: TENANT_ID, CLIENT_ID, or CLIENT_SECRET is not set. "
            "Please set these environment variables."
        )
        return

    # --- New SharePoint Download Example ---
    example_sharepoint_url = "https://1drv.ms/b/c/15b85433fda478fe/EcBJk4nBPZZNiVIxKBCNeRYBAU53LoOdHCgLPESiQ0yn3A?e=NeR4pI"  # Replace this
    example_output_pdf_path = "downloaded_document.pdf"

    print("\n--- Running SharePoint Download Example ---")
    await download_sharepoint_url_as_pdf(
        client, example_sharepoint_url, example_output_pdf_path
    )
    print("-------------------------------------\n")

    # --- Existing User Listing Logic (Commented out for this example) ---
    # try:
    #     print("Fetching all users...")
    #     users_response = await client.users.get()

    #     if users_response and users_response.value:
    #         all_users = users_response.value
    #         next_link = users_response.odata_next_link
    #         # Basic pagination handling (prints first page, notes if more exist)
    #         if next_link:
    #             print(
    #                 f"Found {len(users_response.value)} users on this page. More users might exist (odata_next_link present)."
    #             )
    #             print(
    #                 "To fetch all users across pages, full pagination handling is required."
    #             )

    #         print(f"\nFound {len(all_users)} user(s) on the first page:")
    #         for user in all_users:
    #             print(
    #                 f"  Display Name: {user.display_name}, UserPrincipalName: {user.user_principal_name}, ID: {user.id}"
    #             )
    #     else:
    #         print("No users found or an error occurred while fetching users.")

    # except Exception as e:
    #     print(f"An error occurred while fetching users: {e}")


if __name__ == "__main__":
    asyncio.run(main())
