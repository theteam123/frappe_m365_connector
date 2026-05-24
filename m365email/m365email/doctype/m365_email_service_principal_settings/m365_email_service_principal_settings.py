# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document


class M365EmailServicePrincipalSettings(Document):
	"""Store M365 app registration credentials per tenant (supports multiple tenants)"""
	
	def validate(self):
		"""Validate service principal settings"""
		# NOTE: Fix authority URL if it contains placeholder or is missing
		if self.tenant_id:
			if not self.authority_url or "{tenant_id}" in self.authority_url:
				self.authority_url = f"https://login.microsoftonline.com/{self.tenant_id}"
		
		# Set default Graph API endpoint if not provided
		if not self.graph_api_endpoint:
			self.graph_api_endpoint = "https://graph.microsoft.com/v1.0"
		
		# Set default scopes if not provided
		if not self.scopes:
			self.scopes = "https://graph.microsoft.com/.default"
	
	def on_update(self):
		"""Clear token cache when credentials are updated"""
		if self.has_value_changed("client_id") or self.has_value_changed("client_secret") or self.has_value_changed("tenant_id"):
			self.token_cache = None
			self.last_token_refresh = None
			self.token_expires_at = None

