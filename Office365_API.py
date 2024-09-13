
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File


USERNAME = "ismm21103@student.nitw.ac.in"
PASSWORD = "IMRANsyed@347"
SHAREPOINT_SITE = "https://nitwacin.sharepoint.com/sites/myproject"
SHAREPOINT_SITE_NAME = "myproject"
SHAREPOINT_DOC = "Shared Documents"

class SharePoint:
    def _auth(self):
        conn = ClientContext(SHAREPOINT_SITE).with_credentials(
            UserCredential(
                USERNAME,
                PASSWORD
            )
        )
        return conn
    
    def _get_files_list(self,folder_name):
        conn = self._auth()
        target_folder_url = f'{SHAREPOINT_DOC}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Files","Folders"]).get().execute_query()
        return root_folder.files
    
    def download_file(self, file_name, folder_name):
        conn = self._auth()
        file_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC}/{folder_name}/{file_name}'
        file = File.open_binary(conn, file_url)
        return file.content
    
    def download_files(self, folder_name):
        return self._get_files_list(folder_name)
    

    
