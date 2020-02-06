# -*- coding: utf-8 -*-
"""
This is a skeleton file that can serve as a starting point for a Python
console script. To run this script uncomment the following lines in the
[options.entry_points] section in setup.cfg:

    console_scripts =
         xlsxutils = another_python_rpa.xlsxutils:run

Then run `python setup.py install` which will install the command `xlsxutils`
inside your current environment.
Besides console scripts, the header (i.e. until _logger...) of this file can
also be used as template for Python modules.

Note: This skeleton file can be safely removed if not needed!
"""

import argparse
import sys
import logging
import xlrd

from another_python_rpa import __version__

__author__ = "Arthur Aleksandro Alves Silva"
__copyright__ = "Arthur Aleksandro Alves Silva"
__license__ = "mit"

_logger = logging.getLogger(__name__)


def extract(file, missfield ="MissingNo"):
  """Extrator de informações de planilhas xlsx

  Args:
    file (string): Arquivo
    missfield (string): Texto para por no lugar das células vazias

  Returns:
    dic: dicionário com informações da planilha
  """
  obj = {
    "nrows": 0,
    "rows": []
  }    
  loc = (file)
  wb = xlrd.open_workbook(loc) 
  sheet = wb.sheet_by_index(0)

  if sheet.nrows > 0:
    obj["nrows"] = sheet.nrows
    nameCols = []

    for i in range(sheet.ncols): 
      nameCols.append(sheet.cell_value(0, i) if sheet.cell_value(0, i) else missfield)
    
    for i in range(1, sheet.nrows):
      obj["rows"].append({})
      for x in range(sheet.ncols):
        celVal = sheet.cell_value(i, x)
        obj["rows"][i-1][nameCols[x]] = celVal if celVal else missfield
  
  return obj



def parse_args(args):
  """Parse command line parameters

  Args:
    file (str): command line paramete as file to be open

  Returns:
    :obj:`argparse.Namespace`: command line parameters namespace
  """
  parser = argparse.ArgumentParser(
    description="Utilitário de processamento de planilha")
  parser.add_argument(
    "--version",
    action="version",
    version="another-python-rpa {ver}".format(ver=__version__),
    help="Mostra a versão do script")
  parser.add_argument(
    "--file",
    "-f",
    dest="file",
    help="Arquivo de planilha",
    required=False,
    type=str,
    metavar="FILE")
  parser.add_argument(
    "-mf",
    "--missfield",
    dest="missfield",
    help="Usa para preencher campos da planilha que estiverem vazios. Default: MissingNo",
    default="MissingNo",
    type=str,
    metavar="MISS FIELD")
  parser.add_argument(
    "-v",
    "--verbose",
    dest="loglevel",
    help="set loglevel to INFO",
    action="store_const",
    const=logging.INFO)
  parser.add_argument(
    "-vv",
    "--very-verbose",
    dest="loglevel",
    help="set loglevel to DEBUG",
    action="store_const",
    const=logging.DEBUG)
  return parser.parse_args(args)


def setup_logging(loglevel):
  """Setup basic logging

  Args:
    loglevel (int): minimum loglevel for emitting messages
  """
  logformat = "[%(asctime)s] %(levelname)s:%(name)s:%(message)s"
  logging.basicConfig(level=loglevel, stream=sys.stdout,
                      format=logformat, datefmt="%Y-%m-%d %H:%M:%S")


def main(args):
  """Main entry point allowing external calls

  Args:
    args ([str]): command line parameter list
  """
  args = parse_args(args)
  setup_logging(args.loglevel)
  _logger.debug("Iniciando processamento...")
  file = args.file
  if not file:
    print("Informe o arquivo (caminho completo): ")
    file = input()

  print("O arquivo {} possui as informações abaixo:\n\n {}".format(file, extract(file, args.missfield)))
  _logger.info("Script ends here")


def run():
    """Entry point for console_scripts
    """
    main(sys.argv[1:])


if __name__ == "__main__":
    run()
