"""Neutral moniker fixtures plus Windows structured-storage growth checks."""
import base64
import os
from pathlib import Path
import struct
import sys
import tempfile
import unittest
from unittest.mock import patch

SCRIPTS = Path(__file__).resolve().parents[1] / 'skills/thinkcell-edit/scripts'
sys.path.insert(0, str(SCRIPTS / 'external_links'))
sys.path.insert(0, str(SCRIPTS))
import native_moniker as native
import portable_rebind as rebind
from test_external_links import carrier, synthetic_ole


def file_moniker(path, unicode=True):
    ansi = path.encode('cp1252', 'replace') + b'\0'
    prefix = native.FILE_CLSID + struct.pack('<HI', 0, len(ansi)) + ansi
    prefix += b'\xff\xff\xad\xde' + b'\0' * 20
    encoded = path.encode('utf-16-le')
    return prefix + (struct.pack('<IIH', len(encoded) + 6, len(encoded), 3) + encoded
                     if unicode else struct.pack('<I', 0))


def item_moniker(name, unicode=False):
    delim = b'!\0' + (b'!\0' if unicode else b'')
    item = name.encode('cp1252') + b'\0' + (name.encode('utf-16-le') if unicode else b'')
    return native.ITEM_CLSID + struct.pack('<I', len(delim)) + delim + struct.pack('<I', len(item)) + item


def composite(path, unicode=True):
    return native.COMPOSITE_CLSID + struct.pack('<I', 2) + file_moniker(path, unicode) + item_moniker('Sales!range', unicode)


class StandardMonikerTests(unittest.TestCase):
    def test_arbitrary_lengths_preserve_exact_range_moniker(self):
        for target in (r'C:\x.xlsx', r'D:\long-folder\a-much-longer-workbook.xlsx', 'Z:\\Unicode\\😀.xlsx'):
            with self.subTest(target=target):
                payload = composite(r'C:\source.xlsx')
                xml, old = carrier()
                xml = xml.replace(base64.b64encode(old), base64.b64encode(payload))
                with patch.object(rebind, 'serialize_file', return_value=file_moniker(target)):
                    changed, source = rebind.rebind_xml(xml, payload, target)
                rebound = rebind._fields(changed)[2]
                _, old_end, _ = native.composite_file(payload)
                _, end, actual = native.composite_file(rebound)
                self.assertEqual(actual, target)
                self.assertEqual(rebound[end:], payload[old_end:])
                self.assertEqual(source, r'C:\source.xlsx')

    def test_ansi_only_and_hybrid_identity_readback(self):
        for unicode in (False, True):
            payload = composite(r'C:\source.xlsx', unicode)
            self.assertEqual(rebind._moniker(base64.b64encode(payload).decode())[0],
                             r'C:\source.xlsx;Sales!range')

    def test_corrupt_unicode_extent_rejected_before_serializer(self):
        payload = bytearray(composite(r'C:\source.xlsx'))
        # The mandatory Unicode extent and byte counts must agree.
        start = 20 + 16 + 6 + len(r'C:\source.xlsx'.encode()) + 1 + 24
        struct.pack_into('<I', payload, start, 7)
        with self.assertRaisesRegex(RuntimeError, 'size or key'):
            native.composite_file(bytes(payload))

    def test_extra_item_bytes_and_invalid_delimiter_rejected(self):
        with self.assertRaisesRegex(RuntimeError, 'trailing'):
            native.item_name(item_moniker('Sales!range') + b'x')
        with self.assertRaisesRegex(RuntimeError, 'delimiter'):
            native.item_name(item_moniker('Sales!range').replace(b'!\0', b'?\0', 1))

    @unittest.skipUnless(os.name == 'nt', 'Windows moniker API')
    def test_real_os_serialization_unicode_and_ansi_paths(self):
        for path in (r'C:\neutral.xlsx', 'C:\\Unicode\\😀.xlsx'):
            raw = native.serialize_composite(path, 'Sales!__neutral')
            _, end, actual = native.composite_file(raw)
            self.assertEqual(actual, path)
            self.assertEqual(native.item_name(raw, end), 'Sales!__neutral')

    @unittest.skipUnless(os.name == 'nt', 'Windows structured-storage API')
    def test_grow_and_shrink_real_cfb_preserves_other_streams(self):
        xml, _ = carrier()
        raw = synthetic_ole(xml)
        grown_xml = xml + b' ' * 7000
        grown = rebind.replace_xml_stream(raw, grown_xml)
        self.assertEqual(rebind._xml(grown), grown_xml)
        shrunk = rebind.replace_xml_stream(grown, xml)
        self.assertEqual(rebind._xml(shrunk), xml)


if __name__ == '__main__':
    unittest.main()
