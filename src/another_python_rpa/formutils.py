# -*- coding: utf-8 -*-
"""
This is a skeleton file that can serve as a starting point for a Python
console script. To run this script uncomment the following lines in the
[options.entry_points] section in setup.cfg:

    console_scripts =
         formutils = another_python_rpa.formutils:run

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


def util(url, op):
    """Leitor e Operador de formulários google

    Args:
      url (string): URL do formulário
      op (string): Operação a ser realizada

    Returns:
      str: Informações sobre a operação escolhida
    """
    



def parse_args(args):
    """Parse command line parameters

    Args:
      url (str): command line paramete as file to be open
      op (str): operação a ser realizada
      ans (list): lista de respostas

    Returns:
      :obj:`argparse.Namespace`: command line parameters namespace
    """
    parser = argparse.ArgumentParser(
        description="Utilitário de processamento de formulário")
    parser.add_argument(
        "--version",
        action="version",
        version="another-python-rpa {ver}".format(ver=__version__),
        help="Mostra a versão do script")
    groupExcl = parser.add_mutually_exclusive_group(required=True)
    groupExcl.add_argument(
      "--get",
      dest="op",
      action="store_true")
    groupExcl.add_argument(
      "--post",
      dest="op",
      action="store_false")

    groupPost = parser.add_argument_group("POST")
   
    parser.add_argument(
      dest="url",
      help="URL do formulário",
      type=str,
      metavar="URL")
    
    groupPost.add_argument(
      "--questans",
      "-qa",
      dest="questans",
      help="Questões do formulário com respostas juntas por delimitador",
      nargs="+",
      required='--post' in sys.argv,
      metavar="QUESTS & ANS")
    
    groupPost.add_argument(
      "--delim",
      "-dl",
      dest="delim",
      help="Delimitador, o padrão é '#delim#'",
      type=str)

    parser.set_defaults(delim="#delim#")

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
    print("O formulário {} passará pela operação {}".format(args.url, "GET" if args.op else "POST"))
    _logger.info("Script ends here")


def run():
    """Entry point for console_scripts
    """
    main(sys.argv[1:])


if __name__ == "__main__":
    run()
