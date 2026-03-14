"""
Bulk Email Pro — CLI Tool
Command-line interface for bulk email sending with Rich output.
"""
import sys, os, json, time
from pathlib import Path

# Add parent for shared modules
sys.path.insert(0, str(Path(__file__).parent.parent / 'desktop_app'))

try:
    import click
    from rich.console import Console
    from rich.table import Table
    from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeRemainingColumn
    from rich.panel import Panel
    from rich.text import Text
    RICH_AVAILABLE = True
except ImportError:
    RICH_AVAILABLE = False
    import click  # click is still required

from account_manager import AccountManager
from excel_processor import ExcelProcessor
from smtp_engine import SMTPEngine

DATA_DIR = Path(os.path.expanduser("~")) / ".bulk_email_pro"
DATA_DIR.mkdir(parents=True, exist_ok=True)

console = Console() if RICH_AVAILABLE else None

def get_am():
    return AccountManager(str(DATA_DIR))

def print_msg(msg, style=""):
    if console:
        console.print(msg, style=style)
    else:
        click.echo(msg)

@click.group()
@click.version_option("3.5.0", prog_name="Bulk Email Pro CLI")
def cli():
    """📧 Bulk Email Pro CLI — Send bulk emails from your own accounts."""
    pass

# ─── ACCOUNTS ─────────────────────────────────────────────────
@cli.group()
def accounts():
    """Manage email accounts."""
    pass

@accounts.command("list")
def accounts_list():
    """List all configured email accounts."""
    am = get_am()
    accs = am.get_all_accounts()
    if not accs:
        print_msg("No accounts configured. Use 'accounts add' to add one.", "yellow")
        return
    if RICH_AVAILABLE:
        table = Table(title="📋 Email Accounts", border_style="dim")
        table.add_column("ID", style="dim")
        table.add_column("Nickname", style="bold")
        table.add_column("Email", style="cyan")
        table.add_column("Status")
        table.add_column("Sent/Limit", justify="right")
        for a in accs:
            st_style = "green" if a["status"]=="connected" else "red" if a["status"]=="failed" else "yellow"
            status = f"[{st_style}]{a['status']}[/{st_style}]"
            table.add_row(a["id"], a["nickname"], a["email"], status, f"{a['sent_today']}/{a['daily_limit']}")
        console.print(table)
    else:
        for a in accs:
            click.echo(f"  {a['id']}  {a['nickname']:15s}  {a['email']:30s}  {a['status']:12s}  {a['sent_today']}/{a['daily_limit']}")
    stats = am.get_stats()
    print_msg(f"\n{stats['total_accounts']} accounts | {stats['total_remaining_today']} remaining | Avg Health: {stats.get('avg_health_score',100)}%", "dim")

@accounts.command("stats")
def accounts_stats():
    """Show detailed account health and campaign statistics."""
    am = get_am()
    accs = am.get_all_accounts()
    if not accs:
        print_msg("No accounts configured.", "yellow"); return
    if RICH_AVAILABLE:
        table = Table(title="❤️ Account Health Dashboard", border_style="dim")
        table.add_column("Nickname", style="bold")
        table.add_column("Email", style="cyan")
        table.add_column("Health", justify="right")
        table.add_column("Total Sent", justify="right")
        table.add_column("Bounces", justify="right")
        table.add_column("Status")
        for a in accs:
            h = a.get('health_score', 100)
            h_style = "green" if h > 70 else "yellow" if h > 40 else "red"
            st_style = "green" if a["status"]=="connected" else "red" if a["status"]=="failed" else "yellow"
            table.add_row(a["nickname"], a["email"],
                         f"[{h_style}]{h}%[/{h_style}]",
                         str(a.get('total_sent',0)),
                         str(a.get('hard_bounces',0)),
                         f"[{st_style}]{a['status']}[/{st_style}]")
        console.print(table)
    stats = am.get_stats()
    print_msg(f"\nLifetime Sent: {stats.get('lifetime_sent',0)} | Failed: {stats.get('lifetime_failed',0)}", "dim")

@accounts.command("export")
@click.option("--output", default="accounts_backup.json", type=click.Path(), help="Output file path")
def accounts_export(output):
    """Export all accounts to a backup file."""
    am = get_am()
    if am.export_accounts(output):
        print_msg(f"✅ Exported {len(am.accounts)} accounts to {output}", "bold green")
    else:
        print_msg("❌ Export failed", "red")

@accounts.command("import")
@click.option("--input", "input_file", required=True, type=click.Path(exists=True), help="Backup file to import")
@click.option("--overwrite", is_flag=True, help="Overwrite existing accounts")
def accounts_import(input_file, overwrite):
    """Import accounts from a backup file."""
    am = get_am()
    result = am.import_accounts(input_file, overwrite)
    if result.get("success"):
        print_msg(f"✅ Imported {result['imported']} accounts (skipped {result['skipped']})", "bold green")
    else:
        print_msg(f"❌ Import error: {result.get('error')}", "red")

@accounts.command("add")
@click.option("--nickname", prompt="Account nickname", help="Friendly name")
@click.option("--email", prompt="Email address", help="Your email")
@click.option("--password", prompt=True, hide_input=True, help="Password or app password")
@click.option("--host", prompt="SMTP host", help="e.g. smtp.gmail.com")
@click.option("--port", prompt="SMTP port", default=587, type=int)
@click.option("--security", prompt="Security (TLS/SSL/None)", default="TLS")
@click.option("--limit", default=500, type=int, help="Daily send limit")
def accounts_add(nickname, email, password, host, port, security, limit):
    """Add a new email account."""
    am = get_am()
    try:
        am.add_account(nickname, email, host, port, security, password, limit)
        print_msg(f"✅ Account '{nickname}' added successfully!", "bold green")
    except ValueError as e:
        print_msg(f"❌ Error: {e}", "bold red")

@accounts.command("test")
@click.argument("nickname")
def accounts_test(nickname):
    """Test connection for an account by nickname."""
    am = get_am()
    acc = am.get_account_by_nickname(nickname)
    if not acc:
        print_msg(f"❌ Account '{nickname}' not found", "red")
        return
    print_msg(f"🔌 Testing {acc['email']}...", "yellow")
    result = am.test_account(acc["id"])
    if result["success"]:
        print_msg(f"✅ {result['message']}", "bold green")
    else:
        print_msg(f"❌ {result['message']}", "bold red")

@accounts.command("delete")
@click.argument("nickname")
@click.confirmation_option(prompt="Are you sure?")
def accounts_delete(nickname):
    """Delete an account by nickname."""
    am = get_am()
    acc = am.get_account_by_nickname(nickname)
    if not acc:
        print_msg(f"❌ Account '{nickname}' not found", "red")
        return
    am.delete_account(acc["id"])
    print_msg(f"🗑️ Account '{nickname}' deleted", "yellow")

# ─── VALIDATE ─────────────────────────────────────────────────
@cli.command()
@click.option("--excel", required=True, type=click.Path(exists=True), help="Path to Excel/CSV file")
@click.option("--column", default=None, help="Email column name (auto-detect if omitted)")
def validate(excel, column):
    """Validate emails from an Excel/CSV file without sending."""
    ep = ExcelProcessor()
    result = ep.load(excel)
    if not result["success"]:
        print_msg(f"❌ Failed to load: {result['error']}", "red")
        return
    print_msg(f"📄 Loaded {result['rows']} rows from {os.path.basename(excel)}", "cyan")
    if not column:
        column = ep.auto_detect_email_column()
        if column:
            print_msg(f"🔍 Auto-detected email column: '{column}'", "green")
        else:
            print_msg("❌ Could not auto-detect email column. Use --column", "red")
            return
    stats = ep.validate_and_load(column)
    print_msg(f"\n✅ Valid: {stats['valid']}", "green")
    print_msg(f"❌ Invalid: {stats['invalid']}", "red")
    print_msg(f"🔄 Duplicates: {stats['duplicates']}", "yellow")
    print_msg(f"🔧 Typos fixed: {stats['typos_fixed']}", "cyan")
    if stats.get('disposable', 0) > 0:
        print_msg(f"⚠️  Disposable emails blocked: {stats['disposable']}", "yellow")
    if stats.get('mx_failed', 0) > 0:
        print_msg(f"🚫 MX validation failed: {stats['mx_failed']}", "red")
    # Domain stats
    domain_stats = ep.get_domain_stats()
    if domain_stats and RICH_AVAILABLE:
        dt = Table(title="🌐 Domain Distribution", border_style="dim")
        dt.add_column("Domain"); dt.add_column("Count", justify="right"); dt.add_column("%", justify="right")
        for d, info in list(domain_stats.items())[:8]:
            dt.add_row(d, str(info['count']), f"{info['percent']}%")
        console.print(dt)
    print_msg(f"\nPersonalization variables: {', '.join(ep.get_personalization_vars())}", "dim")

# ─── SEND ─────────────────────────────────────────────────────
@cli.command()
@click.option("--excel", required=True, type=click.Path(exists=True))
@click.option("--column", default=None, help="Email column name")
@click.option("--subject", required=True, help="Email subject line")
@click.option("--body-file", default=None, type=click.Path(exists=True), help="Path to HTML/text body file")
@click.option("--body", default=None, help="Inline email body text")
@click.option("--account", default=None, help="Account nickname to use")
@click.option("--rotate", is_flag=True, help="Use all accounts in rotation")
@click.option("--delay", default=1.5, type=float, help="Delay between emails (seconds)")
@click.option("--limit", "send_limit", default=None, type=int, help="Only send to first N emails")
@click.option("--dry-run", is_flag=True, help="Preview without sending")
@click.option("--output", default=None, type=click.Path(), help="Save report to CSV")
@click.option("--html/--text", "is_html", default=True, help="HTML or plain text format")
@click.option("--attachment", multiple=True, type=click.Path(exists=True), help="Path to file attachment (can be used multiple times)")
def send(excel, column, subject, body_file, body, account, rotate, delay, send_limit, dry_run, output, is_html, attachment):
    """Send a bulk email campaign."""
    # Load body
    if body_file:
        with open(body_file, 'r', encoding='utf-8') as f:
            email_body = f.read()
    elif body:
        email_body = body
    else:
        print_msg("❌ Provide --body or --body-file", "red")
        return

    # Load recipients
    ep = ExcelProcessor()
    result = ep.load(excel)
    if not result["success"]:
        print_msg(f"❌ {result['error']}", "red"); return
    if not column:
        column = ep.auto_detect_email_column()
        if not column:
            print_msg("❌ Cannot auto-detect email column. Use --column", "red"); return
    stats = ep.validate_and_load(column)
    emails = ep.valid_emails
    if send_limit:
        emails = emails[:send_limit]
    print_msg(f"📄 {len(emails)} valid recipients loaded", "cyan")

    # Setup account
    am = get_am()
    if not rotate and account:
        acc = am.get_account_by_nickname(account)
        if not acc:
            print_msg(f"❌ Account '{account}' not found", "red"); return
    elif not am.get_all_accounts():
        print_msg("❌ No accounts configured. Use 'accounts add'", "red"); return

    if dry_run:
        print_msg("\n🔍 DRY RUN — Preview of first 3 emails:", "bold yellow")
        engine = SMTPEngine()
        for i, recip in enumerate(emails[:3]):
            p_subj = engine.personalize(subject, recip)
            p_body = engine.personalize(email_body, recip)
            if RICH_AVAILABLE:
                console.print(Panel(f"To: {recip['_email']}\nSubject: {p_subj}\n\n{p_body[:200]}...",
                                    title=f"Email #{i+1}", border_style="cyan"))
            else:
                click.echo(f"\n--- Email #{i+1} ---\nTo: {recip['_email']}\nSubject: {p_subj}\n{p_body[:200]}...")
        print_msg(f"\n✅ Dry run complete. {len(emails)} emails would be sent.", "green")
        return

    # Send
    engine = SMTPEngine()
    sent_count = 0
    fail_count = 0
    results = []

    if RICH_AVAILABLE:
        with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"),
                      BarColumn(bar_width=40), TextColumn("{task.completed}/{task.total}"),
                      TimeRemainingColumn(), console=console) as progress:
            task = progress.add_task("📤 Sending emails...", total=len(emails))

            def cb(idx, total, result):
                nonlocal sent_count, fail_count
                if result.get("status") == "sent":
                    sent_count += 1
                elif result.get("status") == "failed":
                    fail_count += 1
                progress.update(task, completed=idx,
                               description=f"✅ {sent_count} sent  ❌ {fail_count} failed")

            results = engine.send_bulk(
                email_list=emails, account_manager=am,
                subject_template=subject, body_template=email_body,
                is_html=is_html, delay=delay, use_rotation=rotate,
                attachment_paths=list(attachment),
                progress_callback=cb)
    else:
        def cb(idx, total, result):
            nonlocal sent_count, fail_count
            st = result.get("status", "")
            if st == "sent": sent_count += 1
            elif st == "failed": fail_count += 1
            pct = round(idx/total*100)
            click.echo(f"\r[{pct}%] Sent: {sent_count} Failed: {fail_count} ({idx}/{total})", nl=False)

        results = engine.send_bulk(
            email_list=emails, account_manager=am,
            subject_template=subject, body_template=email_body,
            is_html=is_html, delay=delay, use_rotation=rotate,
            progress_callback=cb)
        click.echo()

    # Summary
    total = len(results)
    sent = sum(1 for r in results if r.get("status") == "sent")
    failed = sum(1 for r in results if r.get("status") == "failed")
    rate = round(sent/max(total,1)*100, 1)

    if RICH_AVAILABLE:
        console.print(Panel(
            f"✅ Sent: {sent}\n❌ Failed: {failed}\n📊 Success Rate: {rate}%",
            title="🎉 Campaign Complete", border_style="green"))
    else:
        click.echo(f"\n🎉 Complete! Sent: {sent}, Failed: {failed}, Rate: {rate}%")

    # Export
    if output:
        import csv
        with open(output, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=["_email","status","account_used","timestamp","error"])
            w.writeheader(); w.writerows(results)
        print_msg(f"📥 Report saved to {output}", "green")

# ─── TEST ─────────────────────────────────────────────────────
@cli.command("test")
@click.argument("nickname")
def test_cmd(nickname):
    """Test SMTP connection for an account."""
    am = get_am()
    acc = am.get_account_by_nickname(nickname)
    if not acc:
        print_msg(f"❌ Account '{nickname}' not found", "red"); return
    print_msg(f"🔌 Testing {acc['email']}...", "yellow")
    result = am.test_account(acc["id"])
    style = "bold green" if result["success"] else "bold red"
    icon = "✅" if result["success"] else "❌"
    print_msg(f"{icon} {result['message']}", style)

if __name__ == "__main__":
    cli()
