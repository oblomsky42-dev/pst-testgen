"""Everything that talks to Outlook: the PST store, Extended MAPI writes, deletions.

Why Extended MAPI: the Outlook object model resets PR_MESSAGE_DELIVERY_TIME to
"now" when it saves a new item - whether the date is set before or after the
first save, on a MailItem or a PostItem (verified with Outlook 2024). Writing
the message through Extended MAPI keeps delivery time, submit time, read state,
sender and Message-ID exactly as given.
"""

from __future__ import annotations

import os
from datetime import datetime, timezone

try:
    import pythoncom
    import pywintypes
    import win32com.client
    from win32com.mapi import mapi
    HAS_WIN32 = True
except ImportError:  # not Windows, or pywin32 missing
    HAS_WIN32 = False

# MAPI property tags (PT_UNICODE 0x001F, PT_SYSTIME 0x0040, PT_LONG 0x0003, PT_BINARY 0x0102)
PR_MESSAGE_CLASS_W = 0x001A001F
PR_SUBJECT_W = 0x0037001F
PR_BODY_W = 0x1000001F
PR_SENDER_NAME_W = 0x0C1A001F
PR_SENDER_EMAIL_ADDRESS_W = 0x0C1F001F
PR_SENDER_ADDRTYPE_W = 0x0C1E001F
PR_SENT_REPRESENTING_NAME_W = 0x0042001F
PR_SENT_REPRESENTING_EMAIL_ADDRESS_W = 0x0065001F
PR_SENT_REPRESENTING_ADDRTYPE_W = 0x0064001F
PR_INTERNET_MESSAGE_ID_W = 0x1035001F
PR_RECIPIENT_TYPE = 0x0C150003
PR_DISPLAY_NAME_W = 0x3001001F
PR_ADDRTYPE_W = 0x3002001F
PR_EMAIL_ADDRESS_W = 0x3003001F
PR_MESSAGE_DELIVERY_TIME = 0x0E060040
PR_CLIENT_SUBMIT_TIME = 0x00390040
PR_MESSAGE_FLAGS = 0x0E070003
PR_ATTACH_METHOD = 0x37050003
PR_ATTACH_FILENAME_W = 0x3704001F
PR_ATTACH_LONG_FILENAME_W = 0x3707001F
PR_ATTACH_DATA_BIN = 0x37010102
MSGFLAG_READ = 0x0001
ATTACH_BY_VALUE = 1
OL_STORE_UNICODE = 3
OL_FOLDER_DELETED_ITEMS = 3


class MapiFolderWriter:
    """Writes received-state messages into one folder of a PST through Extended MAPI."""

    def __init__(self, store_id: str, folder_id: str):
        mapi.MAPIInitialize(None)
        self._session = mapi.MAPILogonEx(
            0, "", None,
            mapi.MAPI_EXTENDED | mapi.MAPI_USE_DEFAULT | mapi.MAPI_NEW_SESSION | mapi.MAPI_NO_MAIL)
        self._store = self._session.OpenMsgStore(
            0, bytes.fromhex(store_id), None,
            mapi.MDB_WRITE | mapi.MDB_NO_MAIL | mapi.MAPI_DEFERRED_ERRORS)
        self._folder = self._store.OpenEntry(
            bytes.fromhex(folder_id), None, mapi.MAPI_MODIFY | mapi.MAPI_DEFERRED_ERRORS)

    def create(self, spec: dict, body: str, attach_path: str | None = None) -> datetime:
        """Write one message from a plan.message_spec(); returns the stored delivery time."""
        when = pywintypes.Time(spec["date"].replace(tzinfo=timezone.utc))
        sender = spec["sender"]
        msg = self._folder.CreateMessage(None, 0)
        try:
            msg.SetProps([
                (PR_MESSAGE_CLASS_W, "IPM.Note"),
                (PR_SUBJECT_W, spec["subject"]),
                (PR_BODY_W, body),
                (PR_MESSAGE_DELIVERY_TIME, when),
                (PR_CLIENT_SUBMIT_TIME, when),
                (PR_MESSAGE_FLAGS, MSGFLAG_READ),  # received and read, not a draft
                (PR_SENDER_NAME_W, sender),
                (PR_SENDER_EMAIL_ADDRESS_W, sender),
                (PR_SENDER_ADDRTYPE_W, "SMTP"),
                (PR_SENT_REPRESENTING_NAME_W, sender),
                (PR_SENT_REPRESENTING_EMAIL_ADDRESS_W, sender),
                (PR_SENT_REPRESENTING_ADDRTYPE_W, "SMTP"),
                (PR_INTERNET_MESSAGE_ID_W, spec["message_id"]),
            ])
            # PR_DISPLAY_TO is computed from the recipient table and cannot be set
            msg.ModifyRecipients(mapi.MODRECIP_ADD, [[
                (PR_RECIPIENT_TYPE, mapi.MAPI_TO),
                (PR_DISPLAY_NAME_W, spec["recipient"]),
                (PR_EMAIL_ADDRESS_W, spec["recipient"]),
                (PR_ADDRTYPE_W, "SMTP"),
            ]])
            if attach_path:
                self._attach(msg, attach_path)
            msg.SaveChanges(0)
            _hr, props = msg.GetProps([PR_MESSAGE_DELIVERY_TIME], 0)
            got = props[0][1]
            return datetime(got.year, got.month, got.day, got.hour, got.minute, got.second)
        finally:
            del msg

    @staticmethod
    def _attach(msg, path: str) -> None:
        _number, attachment = msg.CreateAttach(None, 0)
        name = os.path.basename(path)
        attachment.SetProps([(PR_ATTACH_METHOD, ATTACH_BY_VALUE),
                             (PR_ATTACH_FILENAME_W, name),
                             (PR_ATTACH_LONG_FILENAME_W, name)])
        stream = attachment.OpenProperty(PR_ATTACH_DATA_BIN, pythoncom.IID_IStream, 0,
                                         mapi.MAPI_MODIFY | mapi.MAPI_CREATE)
        with open(path, "rb") as fh:
            while chunk := fh.read(1024 * 1024):
                stream.Write(chunk)
        stream.Commit(0)
        del stream
        attachment.SaveChanges(0)
        del attachment

    def close(self) -> None:
        # every MAPI object must be released before MAPIUninitialize
        self._folder = self._store = None
        if self._session is not None:
            try:
                self._session.Logoff(0, 0, 0)
            finally:
                self._session = None
                mapi.MAPIUninitialize()


class OutlookPst:
    """A new PST attached to the default Outlook profile for the length of a run."""

    def __init__(self, path: str):
        self.path = os.path.abspath(path)
        self.started_here = not _outlook_running()
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.namespace.AddStoreEx(self.path, OL_STORE_UNICODE)
        self.store = next((s for s in self.namespace.Stores if _same_file(s, self.path)), None)
        if self.store is None:
            raise RuntimeError(f"Outlook did not attach {self.path}")
        self.root = self.store.GetRootFolder()

    def folder(self, name: str):
        try:
            return self.root.Folders.Add(name)
        except Exception:
            return self.root.Folders(name)

    def deleted_items(self):
        try:
            return self.store.GetDefaultFolder(OL_FOLDER_DELETED_ITEMS)
        except Exception:
            for i in range(1, self.root.Folders.Count + 1):
                folder = self.root.Folders.Item(i)
                if folder.Name.lower() in ("deleted items", "удаленные", "удалённые"):
                    return folder
        return None

    def close(self) -> None:
        """Detach the PST; quit Outlook if this run started it and nobody is using it.

        Outlook keeps a byte-range lock on a PST until its process exits, and a
        locked PST can be neither read by libpff nor copied.
        """
        try:
            self.namespace.RemoveStore(self.root)
            print("\nPST store detached.")
        except Exception:
            print("\nNOTE: could not detach the store - remove it in Outlook by hand.")
        if not self.started_here:
            print("Outlook was already running: close it before scanning the PST "
                  "(it keeps the file locked until it exits).")
            return
        try:
            if self.outlook.Explorers.Count == 0 and self.outlook.Inspectors.Count == 0:
                self.outlook.Quit()
                print("Outlook (started by this run) closed - the PST is released.")
        except Exception:
            pass


def delete_messages(folder, deleted_items, soft: set, hard: set) -> int:
    """Move soft+hard subjects to Deleted Items, then purge the hard ones. Returns failures."""
    failed = 0
    victims = soft | hard
    items = folder.Items
    for i in range(items.Count, 0, -1):  # backwards: deleting shifts the indices
        try:
            item = items.Item(i)
            if item.Subject in victims:
                item.Delete()  # -> Deleted Items
        except Exception:
            failed += 1
    if hard and deleted_items is not None:
        d_items = deleted_items.Items
        for i in range(d_items.Count, 0, -1):
            try:
                item = d_items.Item(i)
                if item.Subject in hard:
                    item.Delete()  # second delete = permanent (Shift+Delete)
            except Exception:
                failed += 1
    return failed


def _outlook_running() -> bool:
    try:
        win32com.client.GetActiveObject("Outlook.Application")
        return True
    except Exception:
        return False


def _same_file(store, path: str) -> bool:
    try:
        return os.path.normcase(os.path.abspath(store.FilePath)) == os.path.normcase(path)
    except Exception:
        return False
