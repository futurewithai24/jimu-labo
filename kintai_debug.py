"""勤怠連絡メールをOutlook全フォルダから検索するデバッグスクリプト"""
import win32com.client
from datetime import datetime, date

def search_folder(folder, keyword, today, depth=0):
    indent = "  " * depth
    try:
        for item in folder.Items:
            try:
                cls = getattr(item, 'Class', None)
                if cls != 43:
                    continue
                received = getattr(item, 'ReceivedTime', None)
                if not received:
                    continue
                item_date = date(received.year, received.month, received.day)
                if item_date != today:
                    continue
                subject = getattr(item, 'Subject', '') or ''
                if keyword in subject:
                    sender = getattr(item, 'SenderName', '不明')
                    received_str = f"{received.year}/{received.month}/{received.day} {received.hour:02d}:{received.minute:02d}"
                    print(f"{indent}★ 見つかった！")
                    print(f"{indent}  フォルダ: {folder.Name}")
                    print(f"{indent}  件名: {subject}")
                    print(f"{indent}  送信者: {sender}")
                    print(f"{indent}  受信: {received_str}")
            except Exception:
                continue
    except Exception:
        pass

    # サブフォルダも再帰検索
    try:
        for sub in folder.Folders:
            search_folder(sub, keyword, today, depth + 1)
    except Exception:
        pass


outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

today = datetime.now().date()
print(f"「勤怠連絡」を全フォルダから検索中（対象日: {today}）...")
print("=" * 60)

for store in namespace.Stores:
    try:
        root = store.GetRootFolder()
        print(f"\n📁 アカウント: {store.DisplayName}")
        search_folder(root, "勤怠連絡", today)
    except Exception as e:
        print(f"  スキップ: {e}")

print("\n検索完了")
