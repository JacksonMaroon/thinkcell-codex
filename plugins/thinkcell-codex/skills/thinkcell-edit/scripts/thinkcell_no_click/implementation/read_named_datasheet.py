"""Read actual named think-cell datasheet without Office or deck modification."""
from pathlib import Path
import argparse, hashlib, io, json, sys

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1]))
from extract_thinkcell_named_datasheet import extract_named_datasheet


def compact(cells):
    nonempty = [c for c in cells if c['value'] not in (None, '')]
    if not nonempty:
        return {'range': None, 'matrix': [], 'nonempty_cells': []}
    r0, r1 = min(c['row'] for c in nonempty), max(c['row'] for c in nonempty)
    c0, c1 = min(c['column'] for c in nonempty), max(c['column'] for c in nonempty)
    matrix = [[None] * (c1-c0+1) for _ in range(r1-r0+1)]
    for c in nonempty:
        matrix[c['row']-r0][c['column']-c0] = c['value']
    return {'range': {'first_row':r0,'last_row':r1,'first_column':c0,'last_column':c1},
            'matrix':matrix, 'nonempty_cells':nonempty}


def read_blob(data, kind):
    sheets = []
    if kind == 'legacy_biff_cfb':
        import xlrd
        book = xlrd.open_workbook(file_contents=data)
        for sh in book.sheets():
            cells = [{'row':r+1,'column':c+1,'value':sh.cell_value(r,c),'cell_type':sh.cell_type(r,c)}
                     for r in range(sh.nrows) for c in range(sh.ncols)]
            sheets.append({'name':sh.name, **compact(cells)})
    elif kind == 'xlsb_package':
        from pyxlsb import open_workbook, biff12
        from pyxlsb.reader import BIFF12Reader
        from pyxlsb.handlers import CellHandler
        class InlineString(CellHandler):
            def read(self, reader, recid, reclen):
                col, style = reader.read_int(), reader.read_int()
                if recid == 0x3e:
                    reader.skip(1)  # RichStr flags precede XLWideString
                value = reader.read_string()
                if value is None or '\ufffd' in value:
                    raise ValueError('Invalid inline XLSB string')
                return self.cls(col, value, None, style)
        BIFF12Reader.handlers[6] = InlineString()
        BIFF12Reader.handlers[0x3e] = InlineString()
        with open_workbook(io.BytesIO(data)) as book:
            for sn in book.sheets:
                with book.get_sheet(sn) as sh:
                    cells, current_row = [], None
                    # Worksheet.rows filters out 0x3e, so consume records directly.
                    sh._reader.seek(sh._data_offset)
                    for recid, value in sh._reader:
                        if recid == biff12.ROW:
                            current_row = value.r
                        elif biff12.BLANK <= recid <= biff12.FORMULA_BOOLERR or recid == 0x3e:
                            if current_row is None:
                                raise ValueError('Cell before row header')
                            v = value.v
                            if recid == biff12.STRING:
                                if sh._stringtable is None:
                                    raise ValueError('Missing shared string table')
                                v = sh._stringtable[v]
                            cells.append({'row':current_row+1,'column':value.c+1,'value':v,'record_id':recid})
                        elif recid == biff12.SHEETDATA_END:
                            break
                    sheets.append({'name':sn, **compact(cells)})
    else:
        raise ValueError(f'Unsupported storage kind: {kind}')
    return sheets


def read(path, slide, name):
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    data, kind, metadata = extract_named_datasheet(path, slide, name)
    sheets = read_blob(data, kind)
    after = hashlib.sha256(path.read_bytes()).hexdigest()
    if before != after:
        raise RuntimeError('Input changed during read')
    return {'path':str(path),'slide':slide,'automation_name':name,'storage_kind':kind,
            'input_sha256':before,'source_unchanged':True,'metadata':metadata,'sheets':sheets}


if __name__ == '__main__':
    p = argparse.ArgumentParser()
    p.add_argument('path',type=Path)
    p.add_argument('--slide',type=int,required=True)
    p.add_argument('--name',required=True)
    p.add_argument('--output',type=Path)
    a = p.parse_args()
    result = read(a.path.resolve(),a.slide,a.name)
    text = json.dumps(result,indent=2,ensure_ascii=False,allow_nan=False)
    if a.output:
        if a.output.resolve() == a.path.resolve():
            raise ValueError('Output cannot overwrite presentation')
        a.output.write_text(text,encoding='utf-8')
    print(text)
