from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
import os

class SharePointService:
    def __init__(self):
        self.site_url = 'https://uniconsulting079.sharepoint.com/sites/ClearanceFlowAutomation'
        self.ctx = None

    def authenticate_user(self, username, password):
        try:
            auth_context = AuthenticationContext(self.site_url)
            auth_context.acquire_token_for_user(username, password)
            self.ctx = ClientContext(self.site_url, auth_context)
            
            # Verify authentication by making a simple request
            self.ctx.load(self.ctx.web)
            self.ctx.execute_query()
            return True
        except Exception as e:
            # Log error without sensitive information
            print("SharePoint authentication error occurred")
            return False

    def get_context(self):
        if not self.ctx:
            raise Exception("Not authenticated with SharePoint")
        return self.ctx
