import os
import re
import argparse
import logging
import smtplib
import ssl
from typing import List, Optional, Tuple

import pandas as pd
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


def configure_logging(verbose: bool) -> None:
	level = logging.DEBUG if verbose else logging.INFO
	logging.basicConfig(
		level=level,
		format="%(asctime)s %(levelname)s %(message)s",
	)


def parse_args() -> argparse.Namespace:
	parser = argparse.ArgumentParser(
		description="Send invoice PDFs to recipients listed in an Excel sheet"
	)
	parser.add_argument("--excel", required=True, help="Path to Excel file (e.g., data/accounts.xlsx)")
	parser.add_argument("--invoices", required=True, help="Path to invoices folder")
	parser.add_argument("--ext", default=".pdf", help="Invoice file extension to match (default .pdf)")
	parser.add_argument("--dry-run", action="store_true", help="Do not send emails, only log actions")
	parser.add_argument("--subject", help="Override email subject (otherwise from env)")
	parser.add_argument("--body", help="Override email body (otherwise from env)")
	parser.add_argument("--from", dest="from_addr", help="Override From address (otherwise EMAIL_FROM or SMTP_USER)")
	parser.add_argument("--account-column-name", help="Header name for account number column (optional)")
	parser.add_argument("--emails-column-name", help="Header name for recipient emails column (optional)")
	parser.add_argument("--company-column-name", help="Header name for company name column (optional)")
	parser.add_argument("--account-column-index", type=int, default=1, help="Zero-based index for account number column (default 1 for Column B)")
	parser.add_argument("--emails-column-index", type=int, default=6, help="Zero-based index for emails column (default 6 for Column G)")
	parser.add_argument("--company-column-index", type=int, default=0, help="Zero-based index for company name column (default 0 for Column A)")
	parser.add_argument("--verbose", action="store_true", help="Enable debug logging")
	return parser.parse_args()


def load_env() -> None:
	load_dotenv()


def extract_five_digit_account(value: object) -> Optional[str]:
	if value is None:
		return None
	s = str(value).strip()
	if not s:
		return None
	match = re.search(r"\b(\d{5})\b", s)
	if match:
		return match.group(1)
	digits_only = re.sub(r"\D", "", s)
	if len(digits_only) == 5:
		return digits_only
	return None


def split_emails(raw: object) -> List[str]:
	if raw is None:
		return []
	s = str(raw).strip()
	if not s:
		return []
	parts = re.split(r"[;,\s]+", s)
	return [p for p in (part.strip() for part in parts) if p and "@" in p]


def find_invoice_path(account_number: str, invoices_dir: str, ext: str) -> Optional[str]:
	prefix = f"{account_number}_"
	try:
		entries = sorted(os.listdir(invoices_dir))
	except FileNotFoundError:
		logging.error("Invoices directory not found: %s", invoices_dir)
		return None
	for name in entries:
		if not name.lower().endswith(ext.lower()):
			continue
		if name.startswith(prefix):
			full_path = os.path.join(invoices_dir, name)
			if os.path.isfile(full_path):
				return full_path
	return None


def read_excel(excel_path: str) -> pd.DataFrame:
	try:
		df = pd.read_excel(excel_path)
	except Exception as exc:
		logging.error("Failed to read Excel '%s': %s", excel_path, exc)
		raise
	return df


def get_cell_value(row: pd.Series, df_columns: pd.Index, header_name: Optional[str], index_fallback: Optional[int]) -> object:
	if header_name and header_name in df_columns:
		return row[header_name]
	if index_fallback is not None:
		try:
			return row.iloc[index_fallback]
		except Exception:
			return None
	return None


def build_smtp_config_from_env() -> Tuple[str, int, str, str, bool]:
	host = os.getenv("SMTP_HOST", "")
	port_str = os.getenv("SMTP_PORT", "587")
	user = os.getenv("SMTP_USER", "")
	password = os.getenv("SMTP_PASSWORD", "")
	use_tls = os.getenv("SMTP_USE_TLS", "true").lower() in ("1", "true", "yes", "on")
	try:
		port = int(port_str)
	except ValueError:
		port = 587
	return host, port, user, password, use_tls


def send_email_with_attachment(
	from_addr: str,
	to_addrs: List[str],
	subject: str,
	body: str,
	attachment_path: str,
	smtp_host: str,
	smtp_port: int,
	smtp_user: str,
	smtp_password: str,
	use_tls: bool,
) -> None:
	message = MIMEMultipart()
	message["From"] = from_addr
	message["To"] = ", ".join(to_addrs)
	message["Subject"] = subject
	message.attach(MIMEText(body, "plain"))

	with open(attachment_path, "rb") as f:
		part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
	part["Content-Disposition"] = f"attachment; filename=\"{os.path.basename(attachment_path)}\""
	message.attach(part)

	if use_tls:
		context = ssl.create_default_context()
		with smtplib.SMTP(smtp_host, smtp_port) as server:
			server.ehlo()
			server.starttls(context=context)
			server.ehlo()
			if smtp_user:
				server.login(smtp_user, smtp_password)
			server.sendmail(from_addr, to_addrs, message.as_string())
	else:
		with smtplib.SMTP(smtp_host, smtp_port) as server:
			if smtp_user:
				server.login(smtp_user, smtp_password)
			server.sendmail(from_addr, to_addrs, message.as_string())


def smtp_test(smtp_host: str, smtp_port: int, smtp_user: str, smtp_password: str, use_tls: bool) -> Tuple[bool, str]:
	"""Attempt to connect/login to the SMTP server. Returns (ok, message)."""
	try:
		if use_tls:
			context = ssl.create_default_context()
			with smtplib.SMTP(smtp_host, smtp_port, timeout=15) as server:
				server.ehlo()
				server.starttls(context=context)
				server.ehlo()
				if smtp_user:
					server.login(smtp_user, smtp_password)
				return True, "SMTP connection successful"
		else:
			with smtplib.SMTP(smtp_host, smtp_port, timeout=15) as server:
				if smtp_user:
					server.login(smtp_user, smtp_password)
				return True, "SMTP connection successful"
	except Exception as exc:
		return False, f"SMTP test failed: {exc}"


def process_invoices(
	excel_path: str,
	invoices_dir: str,
	from_addr: str,
	subject: str,
	body: str,
	smtp_host: str,
	smtp_port: int,
	smtp_user: str,
	smtp_password: str,
	use_tls: bool,
	ext: str = ".pdf",
	dry_run: bool = True,
	account_column_name: Optional[str] = None,
	emails_column_name: Optional[str] = None,
	company_column_name: Optional[str] = None,
	account_column_index: int = 1,
	emails_column_index: int = 6,
	company_column_index: int = 0,
) -> dict:
	"""Run the invoice sending workflow and return a summary dict."""
	if not os.path.isfile(excel_path):
		raise FileNotFoundError(f"Excel file does not exist: {excel_path}")
	if not os.path.isdir(invoices_dir):
		raise FileNotFoundError(f"Invoices directory does not exist: {invoices_dir}")

	df = read_excel(excel_path)
	processed = sent = skipped = missing_file = 0

	for idx, row in df.iterrows():
		processed += 1
		company = get_cell_value(row, df.columns, company_column_name, company_column_index)
		raw_account = get_cell_value(row, df.columns, account_column_name, account_column_index)
		raw_emails = get_cell_value(row, df.columns, emails_column_name, emails_column_index)

		account = extract_five_digit_account(raw_account)
		recipients = split_emails(raw_emails)

		if not account:
			logging.warning("Row %s: missing/invalid account number; skipping", idx + 1)
			skipped += 1
			continue
		if not recipients:
			logging.warning("Row %s (acct %s): no recipient emails; skipping", idx + 1, account)
			skipped += 1
			continue

		invoice_path = find_invoice_path(account, invoices_dir, ext)
		if not invoice_path:
			logging.warning("Row %s (acct %s): no invoice found in %s", idx + 1, account, invoices_dir)
			missing_file += 1
			continue

		descriptor = f"acct {account}{' - ' + str(company) if pd.notna(company) else ''}".strip()
		if dry_run:
			logging.info("DRY RUN would send %s to %s with attachment %s", descriptor, ", ".join(recipients), os.path.basename(invoice_path))
			continue

		try:
			send_email_with_attachment(
				from_addr=from_addr,
				to_addrs=recipients,
				subject=subject,
				body=body,
				attachment_path=invoice_path,
				smtp_host=smtp_host,
				smtp_port=smtp_port,
				smtp_user=smtp_user,
				smtp_password=smtp_password,
				use_tls=use_tls,
			)
			sent += 1
			logging.info("Sent %s to %s", descriptor, ", ".join(recipients))
		except Exception as exc:
			logging.error("Failed sending %s: %s", descriptor, exc)
			skipped += 1

	return {
		"processed": processed,
		"sent": sent,
		"skipped": skipped,
		"missing_file": missing_file,
	}


def main() -> None:
	args = parse_args()
	configure_logging(args.verbose)
	load_env()

	subject_default = os.getenv("EMAIL_SUBJECT", "Your Monthly Invoice")
	body_default = os.getenv("EMAIL_BODY", "Hello,\n\nPlease find your monthly invoice attached.\n\nThank you.")
	from_default = os.getenv("EMAIL_FROM") or os.getenv("SMTP_USER", "")

	subject = args.subject or subject_default
	body = args.body or body_default
	from_addr = args.from_addr or from_default

	smtp_host, smtp_port, smtp_user, smtp_password, use_tls = build_smtp_config_from_env()

	if not from_addr:
		logging.error("No From address set. Provide --from or EMAIL_FROM/SMTP_USER in env.")
		return
	if not smtp_host:
		logging.error("SMTP_HOST not set in env.")
		return

	try:
		result = process_invoices(
			excel_path=args.excel,
			invoices_dir=args.invoices,
			from_addr=from_addr,
			subject=subject,
			body=body,
			smtp_host=smtp_host,
			smtp_port=smtp_port,
			smtp_user=smtp_user,
			smtp_password=smtp_password,
			use_tls=use_tls,
			ext=args.ext,
			dry_run=args.dry_run,
			account_column_name=args.account_column_name,
			emails_column_name=args.emails_column_name,
			company_column_name=args.company_column_name,
			account_column_index=args.account_column_index,
			emails_column_index=args.emails_column_index,
			company_column_index=args.company_column_index,
		)
		logging.info(
			"Done. processed=%d sent=%d skipped=%d missing_file=%d",
			result["processed"], result["sent"], result["skipped"], result["missing_file"],
		)
	except Exception as exc:
		logging.error("Run failed: %s", exc)


if __name__ == "__main__":
	main()


