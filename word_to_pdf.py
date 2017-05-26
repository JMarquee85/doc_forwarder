import sys
import os
import comtypes.client

import email_pupculture as e_pc
import private_config as p_con
import get_reg as gr
import main_loop
import unoconv
import pywin32
import xtopdf

wdFormatPDF = 17

def convx_to_pdf(infile, outfile):
	"""Converts word doc to pdf"""
	word = comtypes.client.CreateObject('Word.Application')
	doc = word.Documents.Open(infile)
	doc.SaveAs(outfile, FileFormat=wdFormatPDF)
	doc.Close()
	word.Quit()
