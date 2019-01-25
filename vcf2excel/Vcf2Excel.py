"""
This script defines a class to manage the conversion from VCF to Excel 
spreadsheets
"""

from cyvcf2 import VCF
import numpy as np
import pandas as pd
import re
import sys

class Vcf2Excel:
    """
    A class to store and manage a vcf to Excel spreadsheet conversion
    """
    def __init__(self, vcfpath, outpath):
        """
        Instantiate Vcf2Excel object: save input and start parsing

            vcfpath (str): path to a .vcf file

            returns: None
        """
        self.vcfpath = vcfpath
        self.outpath = outpath
        self.vcf = VCF(self.vcfpath)
        self.md_keypairs = pd.DataFrame(columns=['Name', 'Value'])
        self.md_info = pd.DataFrame(columns=['ID', 'Number', 'Type', 
                                                'Description', 'Source',
                                                'Version'])
        self.md_filter = pd.DataFrame(columns=['ID', 'Description'])
        self.md_format = pd.DataFrame(columns=['ID', 'Number', 'Type',
                                                'Description'])
        self.md_alt = pd.DataFrame(columns=['ID', 'Description'])
        self.md_contig = pd.DataFrame(columns=['ID', 'URL'])
        self.md_sample = pd.DataFrame(columns=['ID', 'Genomes', 'Mixture',
                                                'Description'])
        self.md_pedigree = pd.DataFrame(columns=['Name', 'Genome'])

        self.parse_headers()

        self.write_spreadsheet()

    """
    Parse/Tokenize Headers
    """

    def parse_headers(self):
        """
        Builds pandas DataFrames for the different types of possible metadata,
        each with their own parsing function below

            returns: None
        """
        for line in self.vcf.raw_header.split('\n'):
            if line[:2] != "##":
                return

            line = line.strip("##")

            if line.startswith("INFO"):
                self.md_info = self.md_info.append(
                    self.multi_key_parse(line, list(self.md_info.columns)),
                    ignore_index=True)

            elif line.startswith("FILTER"):
                self.md_filter = self.md_filter.append(
                    self.multi_key_parse(line, list(self.md_filter.columns)),
                    ignore_index=True)

            elif line.startswith("FORMAT"):
                self.md_format = self.md_format.append(
                    self.multi_key_parse(line, list(self.md_format.columns)),
                    ignore_index=True)

            elif line.startswith("ALT"):
                self.md_alt = self.md_alt.append(
                    self.multi_key_parse(line, list(self.md_alt.columns)),
                    ignore_index=True)

            elif line.startswith("contig"):
                self.md_contig = self.md_contig.append(
                    self.multi_key_parse(line, list(self.md_contig.columns)),
                    ignore_index=True)

            elif line.startswith("SAMPLE"):
                self.md_sample = self.md_sample.append(
                    self.multi_key_parse(line, list(self.md_sample.columns)),
                    ignore_index=True)

            elif line.startswith("PEDIGREE"):
                self.parse_pedigree(line)

            elif len(line.split('=')) == 2:
                self.parse_keypair(line)

            else:
                raise ValueError("Invalid header line: "+line)

    def parse_keypair(self, line):
        """
        Parses a line of metadata that encodes a single key=value pair

            line (str): key=value format metadata line
            
            returns: None
        """
        key, value = line.split('=')
        self.md_keypairs = self.md_keypairs.append({'Name': key, 
                                                    'Value': value},
                                ignore_index=True)

    def multi_key_parse(self, line, fields):
        """
        Returns a dict of key=value tokens for lines of type INFO, FORMAT,
        FILTER, ALT, contig, or SAMPLE

            line (str): A single metadata line from one of the above
                        categories

            returns (dict): Line parsed into key=value pairs
        """
        line = line.split('=<')[1][:-1]
        pairs = re.findall(r'(?:[^\s,"]|"(?:\\.|[^"])*")+', line)
        keyvals = [pair.split('=') for pair in pairs]
        vals_d = {pair[0]: pair[1] for pair in keyvals}
        new_rec = {}
        for field in fields:
            try:
                new_rec[field] = vals_d[field]

            except KeyError:
                new_rec[field] = np.nan

        return new_rec

    def parse_pedigree(self, line):
        """
        Parses a line of metadata that encodes a PEDIGREE line as described
        by the VCF 4.2 specification

            line (str): A single PEDIGREE metadata line

            returns: None
        """
        raise NotImplementedError

    """
    Output data to spreadsheet
    """

    def write_spreadsheet(self):
        """
        Builds a spreadsheet object and writes it to self.outpath

            returns: None
        """
        writer = pd.ExcelWriter(self.outpath, engine='xlsxwriter')

        self.md_keypairs.to_excel(writer, sheet_name='File Metadata')
        self.md_info.to_excel(writer, sheet_name='INFO')
        self.md_filter.to_excel(writer, sheet_name='FILTER')
        self.md_format.to_excel(writer, sheet_name='FORMAT')
        self.md_alt.to_excel(writer, sheet_name='ALT')
        self.md_contig.to_excel(writer, sheet_name='contig')
        self.md_sample.to_excel(writer, sheet_name='SAMPLE')
        self.md_pedigree.to_excel(writer, sheet_name='PEDIGREE')

        writer.save()


if __name__ == "__main__":
    try:
        vcfpath = sys.argv[1]
        outpath = sys.argv[2]
    except IndexError:
        print("Usage: python Vcf2Excel.py <input_vcf_path> <output_vcf_path>")
        sys.exit(1)
    v2e = Vcf2Excel(vcfpath, outpath)