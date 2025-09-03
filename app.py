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
		invoices_dir = st.text_input("Invoices folder", value=os.getenv("UI_INVOICES", "invoices"))
		ext = st.text_input("Invoice file extension", value=os.getenv("UI_EXT", ".pdf"))

	with st.expander("Email settings", expanded=True):
		default_subject = os.getenv("EMAIL_SUBJECT", "Your Monthly Invoice")
		default_body = os.getenv("EMAIL_BODY", "Hello,\n\nPlease find your monthly invoice attached.\n\nThank you.")
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

	with st.expander("Settings persistence", expanded=False):
		st.caption("Optionally save these values into .env so they load next time.")
		env_path = st.text_input(".env path", value=os.getenv("ENV_PATH", ".env"))
		if st.button("Save settings (.env)"):
			try:
				save_env_file(
					env_path,
					{
						"UI_EXCEL": excel_path,
						"UI_INVOICES": invoices_dir,
						"UI_EXT": ext,
						"SMTP_HOST": smtp_host,
						"SMTP_PORT": int(smtp_port),
						"SMTP_USER": smtp_user,
						"SMTP_PASSWORD": smtp_password,
						"SMTP_USE_TLS": bool(use_tls),
						"EMAIL_FROM": from_addr,
						"EMAIL_SUBJECT": subject,
						"EMAIL_BODY": body,
					},
				)
				st.success(f"Saved settings to {env_path}")
			except Exception as exc:
				st.error(str(exc))

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
				)
				st.success(
					f"Send complete: processed={result['processed']} sent={result['sent']} skipped={result['skipped']} missing_file={result['missing_file']}"
				)
			except Exception as exc:
				st.error(str(exc))


if __name__ == "__main__":
	main()


