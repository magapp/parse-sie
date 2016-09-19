#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
#import os
import argparse
import shlex


def main():
    parser = argparse.ArgumentParser(description='Parse SIE file format')
    parser.add_argument('--filename', type = argparse.FileType('r'), required = True)
    parser.add_argument('--encoding', default = 'cp850', choices = ['cp850', 'latin-1', 'utf-8', 'windows-1252'])
    parser.add_argument('--output', required = False, help = 'Output CSV filename')

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
    attribute_fnamn = "Okänd"
    attribute_gen = ""
    attribute_valuta = "SEK"
    attribute_accounts = dict()
    attribute_dims = dict()
    attribute_kst = dict()
    attribute_proj = dict()
    verifications = list()

    output_file = None
    if args.output:
        output_file = open(args.output, 'w')

    # print "Analyserar..."
    content = args.filename.readlines()
    for l in range(0,len(content)):
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
            l, vers = parse_ver(content, l, args.encoding, vertext, verdate)
            verifications.extend(vers)
    
    # print "Resultat '%s' gen: %s: " % (attribute_fnamn, attribute_gen)

    print '"Date","DateMonth","Account","Account group","Account name","Kst","Proj","Amount","Text"'
    if output_file:
        output_file.write('"Date","DateMonth","Account","Account group","Account name","Kst","Proj","Amount","Text"\n')
    for ver in verifications:
        account_name = ""
        proj_name = ver["proj_nr"]
        kst_name = ver["kst"]
        if attribute_accounts.has_key(ver["account"]):
            account_name = attribute_accounts[ver["account"]]
        if attribute_kst.has_key(ver["kst"]):
            kst_name = attribute_kst[ver["kst"]]
        if attribute_proj.has_key(ver["proj_nr"]):
            proj_name = attribute_proj[ver["proj_nr"]]
        print '"%s-%s-%s","%s-%s","%s","%s","%s","%s","%s","%s","%s"' % (
                    ver["verdate"][0:4], 
                    ver["verdate"][4:6], 
                    ver["verdate"][6:8], 
                    ver["verdate"][0:4], 
                    ver["verdate"][4:6], 
                    ver["account"], 
                    account_group[ver["account"][0]],
                    account_name, 
                    kst_name, 
                    proj_name, 
                    ver["amount"], 
                    ver["vertext"]
             )
        if output_file:
            output_file.write('"%s-%s-%s","%s-%s","%s","%s","%s","%s","%s","%s","%s"\n' % (
                        ver["verdate"][0:4], 
                        ver["verdate"][4:6], 
                        ver["verdate"][6:8], 
                        ver["verdate"][0:4], 
                        ver["verdate"][4:6], 
                        ver["account"], 
                        account_group[ver["account"][0]],
                        account_name, 
                        kst_name, 
                        proj_name, 
                        ver["amount"], 
                        ver["vertext"]
                 ))
    if output_file:
        output_file.close()

def parse(line, encoding):
    if not line.startswith('#') or len(line.split(' ')) == 0:
        return False, False, False
    parts = map(lambda s: s.decode(encoding).encode('utf8'), shlex.split(line.replace('{', '"').replace('}', '"')))
    label = parts[0].upper()[1:]
    text =  ' '.join(parts[1:])
    return label, text, parts[1:]

def parse_ver(content, line, encoding, default_vertext, default_verdate):
    vers = list()
    if content[line + 1].startswith('{'):
        line = line + 2
        while not content[line].startswith('}'):
            label, text, parts = parse(content[line].strip(), encoding)
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

            vers.append({'account': account, 'kst': kst, 'proj_nr': proj, 'amount': amount, 'verdate': verdate, 'vertext': vertext})
            line = line + 1

    return line, vers

if __name__ == "__main__":
    main()
