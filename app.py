import os
import streamlit as st
from dotenv import load_dotenv

from send_invoices import (
	build_smtp_config_from_env,
	process_invoices,
	smtp_test,
)


def _serialize_env_value(value: object) -> str:
	if isinstance(value, bool):
		return "true" if value else "false"
	return str(value)


def save_env_file(env_path: str, values: dict) -> None:
	# Basic .env writer (overwrites file)
	lines = []
	for key, raw in values.items():
		val = _serialize_env_value(raw)
		val = val.replace("\n", "\\n")
		lines.append(f"{key}={val}")
	content = "\n".join(lines) + "\n"
	with open(env_path, "w", encoding="utf-8") as f:
		f.write(content)


def main() -> None:
	st.set_page_config(page_title="Invoice Sender", page_icon="📧", layout="centered")
	st.title("Invoice Sender")
	st.caption("Email invoice PDFs based on an Excel sheet")

	load_dotenv()

	with st.expander("Paths and files", expanded=True):
		excel_path = st.text_input("Excel file", value=os.getenv("UI_EXCEL", "data/accounts.xlsx"))
		sheet_name = st.text_input("Excel sheet name (optional)", value=os.getenv("UI_SHEET", ""), help="Leave blank for first sheet, or specify sheet name like 'Combined'")
		
		# Invoices folder path with better instructions
		st.write("**Invoices folder:**")
		invoices_dir = st.text_input("Invoices folder", value=os.getenv("UI_INVOICES", "invoices"), label_visibility="collapsed", help="Full path to folder containing PDF invoices")
		
		# Helpful instructions for folder selection
		st.info("💡 **To find your folder path:**\n"
				"• **Mac**: Right-click folder → 'Copy as Pathname' or drag folder to Terminal\n"
				"• **Windows**: Right-click folder → 'Copy as path' or hold Shift+right-click → 'Copy as path'\n"
				"• **Example**: `/Users/username/Documents/invoices` or `C:\\Users\\username\\Documents\\invoices`")
		
		ext = st.text_input("Invoice file extension", value=os.getenv("UI_EXT", ".pdf"))

	with st.expander("Excel column settings", expanded=True):
		st.caption("Configure which Excel columns contain your data")
		col1, col2, col3 = st.columns(3)
		with col1:
			account_column_letter = st.selectbox("Account column", options=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'], index=int(os.getenv("ACCOUNT_COLUMN", "1")), help="Column containing 5-digit account numbers")
			account_column_index = ord(account_column_letter) - ord('A')
		with col2:
			emails_column_letter = st.selectbox("Email column", options=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'], index=int(os.getenv("EMAILS_COLUMN", "6")), help="Column containing email addresses")
			emails_column_index = ord(emails_column_letter) - ord('A')
		with col3:
			company_column_letter = st.selectbox("Company column", options=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'], index=int(os.getenv("COMPANY_COLUMN", "0")), help="Column containing company names")
			company_column_index = ord(company_column_letter) - ord('A')

	with st.expander("Email settings", expanded=True):
		default_subject = os.getenv("EMAIL_SUBJECT", "Your Invoice")
		default_body = os.getenv("EMAIL_BODY", "Hello %COMPANY%,\n\nHere is the invoice for account %ACCOUNT%.\n\nThank you.")
		default_from = os.getenv("EMAIL_FROM") or os.getenv("SMTP_USER", "")

		from_addr = st.text_input("From address", value=default_from)
		subject = st.text_input("Subject", value=default_subject)
		body = st.text_area("Body", value=default_body, height=150)

	with st.expander("SMTP server", expanded=True):
		smtp_host_env, smtp_port_env, smtp_user_env, smtp_password_env, use_tls_env = build_smtp_config_from_env()
		smtp_host = st.text_input("SMTP host", value=smtp_host_env)
		smtp_port = st.number_input("SMTP port", value=int(smtp_port_env or 587), step=1)
		smtp_user = st.text_input("SMTP user", value=smtp_user_env)
		smtp_password = st.text_input("SMTP password", value=smtp_password_env, type="password")
		use_tls = st.checkbox("Use TLS", value=use_tls_env)
		
		st.caption("Rate limiting settings (for Office 365 and large batches)")
		delay_between_emails = st.number_input("Delay between emails (seconds)", value=2.1, min_value=0.1, max_value=10.0, step=0.1, help="Add delay between emails to avoid rate limiting (2.1s recommended for Office 365)")
		max_retries = st.number_input("Max retries for failed emails", value=3, min_value=1, max_value=10, step=1, help="Number of retry attempts for failed emails")

	# Save settings section - more prominent
	st.divider()
	st.write("💾 **Save Your Settings**")
	st.caption("Save these settings so they load automatically next time")
	
	col_save1, col_save2 = st.columns([3, 1])
	with col_save1:
		env_path = st.text_input("Save to file", value=".env", help="File to save settings (usually .env)")
	with col_save2:
		if st.button("💾 Save Settings", use_container_width=True, type="primary"):
			try:
				save_env_file(
					env_path,
					{
						"UI_EXCEL": excel_path,
						"UI_SHEET": sheet_name,
						"UI_INVOICES": invoices_dir,
						"UI_EXT": ext,
						"ACCOUNT_COLUMN": account_column_index,
						"EMAILS_COLUMN": emails_column_index,
						"COMPANY_COLUMN": company_column_index,
						"SMTP_HOST": smtp_host,
						"SMTP_PORT": int(smtp_port),
						"SMTP_USER": smtp_user,
						# Note: SMTP_PASSWORD not saved for security
						"SMTP_USE_TLS": bool(use_tls),
						"EMAIL_FROM": from_addr,
						"EMAIL_SUBJECT": subject,
						"EMAIL_BODY": body,
					},
				)
				st.success(f"✅ Settings saved to {env_path}")
			except Exception as exc:
				st.error(f"❌ Error saving: {exc}")

	st.divider()
	col1, col2, col3 = st.columns(3)
	with col1:
		if st.button("Test SMTP", use_container_width=True):
			ok, msg = smtp_test(smtp_host, int(smtp_port), smtp_user, smtp_password, use_tls)
			if ok:
				st.success(msg)
			else:
				st.error(msg)
	with col2:
		if st.button("Dry Run", use_container_width=True):
			try:
				result = process_invoices(
					excel_path=excel_path,
					invoices_dir=invoices_dir,
					from_addr=from_addr,
					subject=subject,
					body=body,
					smtp_host=smtp_host,
					smtp_port=int(smtp_port),
					smtp_user=smtp_user,
					smtp_password=smtp_password,
					use_tls=use_tls,
					ext=ext,
					dry_run=True,
					delay_between_emails=delay_between_emails,
					max_retries=max_retries,
					account_column_index=account_column_index,
					emails_column_index=emails_column_index,
					company_column_index=company_column_index,
					sheet_name=sheet_name if sheet_name else None,
				)
				st.info(f"Dry run complete: processed={result['processed']} missing_file={result['missing_file']}")
			except Exception as exc:
				st.error(str(exc))
	with col3:
		if st.button("Send", use_container_width=True):
			try:
				result = process_invoices(
					excel_path=excel_path,
					invoices_dir=invoices_dir,
					from_addr=from_addr,
					subject=subject,
					body=body,
					smtp_host=smtp_host,
					smtp_port=int(smtp_port),
					smtp_user=smtp_user,
					smtp_password=smtp_password,
					use_tls=use_tls,
					ext=ext,
					dry_run=False,
					delay_between_emails=delay_between_emails,
					max_retries=max_retries,
					account_column_index=account_column_index,
					emails_column_index=emails_column_index,
					company_column_index=company_column_index,
					sheet_name=sheet_name if sheet_name else None,
				)
				st.success(
					f"Send complete: processed={result['processed']} sent={result['sent']} skipped={result['skipped']} missing_file={result['missing_file']}"
				)
			except Exception as exc:
				st.error(str(exc))


if __name__ == "__main__":
	main()


