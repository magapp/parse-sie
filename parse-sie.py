#!/usr/bin/env python
# -*- coding: utf-8 -*-
#import os
import argparse
import shlex
import datetime


def main():
    parser = argparse.ArgumentParser(description='Parse SIE file format')
    parser.add_argument('--filename',
                        type=argparse.FileType('r'),
                        required=True, nargs='+')
    parser.add_argument('--encoding', default='cp850',
                        choices=['cp850', 'latin-1', 'utf-8', 'windows-1252'])
    parser.add_argument('--output', required=False, help='Output CSV filename')
    parser.add_argument('--debug', action='store_true', default=False,
                        help='Display more debug')
    parser.add_argument('--googlesheets', required=False,
                        help='Name of Google Sheets to create. Need creds.json. Note that the spreadsheet must exists and the content may be overwritten.')

    args = parser.parse_args()

    account_group = {
        '1': '1 - Tillgångar',
        '2': '2 - Eget kapital och skulder',
        '3': '3 - Rörelsens inkomster och intäkter',
        '4': '4 - Utgifter och kostnader förädling',
        '5': '5 - Övriga externa rörelseutgifter och kostnader',
        '6': '6 - Övriga externa rörelseutgifter och kostnader',
        '7': '7- Utgifter och kostnader för personal, avskrivningar mm.',
        '8': '8 - Finansiella och andra inkomster/intäkter och utgifter/kostnader'
    }

    output_file = None
    if args.output:
        output_file = open(args.output, 'w')
    data = list()
    data.append(list(["Date", "DateMonth", "Account", "Account group",
                      "Account name", "Kst", "Proj", "Amount", "Text",
                      "Verification", "DateKey", "Company"]))

    if args.debug:
        print('"' + '","'.join(data[0]) + '"')
    if output_file:
        output_file.write('"' + '","'.join(data[0]) + '"\n')
    
    for inputfile in args.filename:
        attribute_fnamn = "Okänd"
        attribute_gen = ""
        attribute_valuta = "SEK"
        attribute_accounts = dict()
        attribute_dims = dict()
        attribute_kst = dict()
        attribute_proj = dict()
        verifications = list()

        content = inputfile.readlines()
        for l in range(0, len(content)):
            line = content[l]
            if not line.startswith('#') or len(line.split(' ')) == 0:
                continue
            label, text, parts = parse(line, args.encoding)

            if label == "VALUTA":
                attribute_valuta = parts[0]
            if label == "GEN":
                attribute_gen = parts[0]
            if label == "FNAMN":
                attribute_fnamn = parts[0]
            if label == "KONTO":
                attribute_accounts[parts[0]] = parts[1]
            if label == "DIM":
                attribute_dims[parts[0]] = parts[1]
            if label == "OBJEKT":
                if parts[0] == '1':
                    attribute_kst[parts[1]] = parts[2]
                if parts[0] == '6':
                    attribute_proj[parts[1]] = parts[2]
            if label == "VER":
                series = parts[0]
                verno = parts[1]
                verdate = parts[2]
                if len(parts) > 3:
                    vertext = parts[3]
                else:
                    vertext = ""
                l, vers = parse_ver(content, l, args.encoding,
                                    vertext, verdate, verno)
                verifications.extend(vers)
        
        # print "Resultat '%s' gen: %s: " % (attribute_fnamn, attribute_gen)
    
        for ver in verifications:
            account_name = ""
            proj_name = ver["proj_nr"]
            kst_name = ver["kst"]
            if ver["account"] in attribute_accounts:
                account_name = attribute_accounts[ver["account"]]
            if ver["kst"] in attribute_kst:
                kst_name = attribute_kst[ver["kst"]]
            if ver["proj_nr"] in attribute_proj:
                proj_name = attribute_proj[ver["proj_nr"]]

            cols = list()
            cols.append("%s-%s-%s" % (ver["verdate"][0:4], ver["verdate"][4:6], ver["verdate"][6:8]))
            cols.append("%s-%s" % (ver["verdate"][0:4], ver["verdate"][4:6]))
            cols.append("%s" % ver["account"])
            if ver["account"][0] in account_group:
                cols.append("%s" % account_group[ver["account"][0]])
            else:
                cols.append("")
            cols.append("%s" % account_name)
            cols.append("%s" % kst_name)
            cols.append("%s" % proj_name)
            cols.append("%0.0f" % float(ver["amount"]))
            cols.append("%s" % ver["vertext"])
            cols.append("%s" % ver["verno"])
            cols.append("%s" % datetime.datetime.strptime(cols[0], "%Y-%m-%d").strftime("%s"))
            cols.append("%s" % attribute_fnamn)
            if args.debug:
                print('"' + '","'.join(cols) + '"')
            if output_file:
                output_file.write('"' + '","'.join(cols) + '"\n')

            data.append(cols)

    if output_file:
        output_file.close()

    if args.googlesheets:
        import gspread
        from oauth2client.service_account import ServiceAccountCredentials
        nbr_rows = len(data)
        nbr_cols = len(data[0])
        print("Google Sheet: " + args.googlesheets)
        scope = ['https://spreadsheets.google.com/feeds']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
        gc = gspread.authorize(credentials)
        sheet = gc.open(args.googlesheets)
        wks = sheet.get_worksheet(0)
       
        print("Resizing to %d rows and %d cols... (range A1:%s%d)" % (nbr_rows, nbr_cols, chr(64 + nbr_cols), nbr_rows))
        wks.resize(rows=nbr_rows, cols=nbr_cols)

        bulk = 1000 
        for i in range(1, nbr_rows + 1, bulk):
            j = i + bulk - 1
            if j > nbr_rows:
                j = nbr_rows
            print("Updating A%d:%s%d ..." % (i, chr(64 + nbr_cols), j))
            cell_list = wks.range('A%d:%s%d' % (i, chr(64 + nbr_cols), j))

            #print "Populating with data..."
            for cell in cell_list:
                if args.debug:
                    print("getting row %d col %d: '%s'" % (cell.row - 1, cell.col - 1, data[cell.row - 1][cell.col - 1]))
                cell.value = data[cell.row - 1][cell.col - 1].replace('å', 'a').replace('ä', 'a').replace('ö', 'o').replace('Å', 'A').replace('Ä', 'A').replace('Ö', 'O')
    
            #print "Updating sheet..."
            wks.update_cells(cell_list)

        """
        for row_nbr, rowdata in enumerate(data):
            print "row %d" % row_nbr
            for col_nbr, coldata in enumerate(rowdata):
                wks.update_cell(row_nbr+1, col_nbr+1, coldata.decode("utf8"))
        """


def parse(line, encoding):
    if not line.startswith('#') or len(line.split(' ')) == 0:
        return False, False, False
    parts = [s for s in shlex.split(line.replace('{', '"').replace('}', '"'))]
    label = parts[0].upper()[1:]
    text = ' '.join(parts[1:])
    return label, text, parts[1:]


def parse_ver(content, line, encoding, default_vertext, default_verdate,
              verno):
    vers = list()
    if content[line + 1].startswith('{'):
        line = line + 2
        while not content[line].startswith('}'):
            label, text, parts = parse(content[line].strip(), encoding)
            #print("label: %s text: %s parts: %s" % (label, text, parts))
            kst = ""
            proj = ""
            account = parts[0]
            p = parts[1]
            amount = parts[2]
            if len(parts) > 3:
                verdate = parts[3]
            else:
                verdate = default_verdate
            if len(parts) > 4:
                vertext = parts[4]
            else:
                vertext = default_vertext

            if len(p.split(' ')) > 0 and p.split(' ')[0] == '1':
                kst = p.split(' ')[1]

            if len(p.split(' ')) > 2 and p.split(' ')[2] == '6':
                proj = p.split(' ')[3]

            vers.append({'account': account, 'kst': kst, 'proj_nr': proj, 'amount': amount, 'verdate': verdate, 'vertext': vertext, 'verno': verno})
            line = line + 1

    return line, vers


if __name__ == "__main__":
    main()
