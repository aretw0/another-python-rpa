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
import asyncio
import urllib.request
from bs4 import BeautifulSoup
import requests

from another_python_rpa import __version__

__author__ = "Arthur Aleksandro Alves Silva"
__copyright__ = "Arthur Aleksandro Alves Silva"
__license__ = "mit"

_logger = logging.getLogger(__name__)

def post_answers(url, questions, answers, verbose=False):
  """Submete respostas a um formulário

  Args:
    url (string): URL do formulário
    questions (list): Questões do formulário no formato 'entry.99999[*]
    answers (list): Respostas do formulário
    verbose (boolean): Verbosidade do método

  Returns:
    int: Códido HTTP da submissão
  """
  if not len(questions) == len(answers):
    raise Exception("A quantidade de questões e respostas não é a mesma")
  
  submit_url = url.replace('/viewform', '/formResponse')
  form_data = {'draftResponse':[],
              'pageHistory':0}
  for q, a in zip(questions, answers):
    form_data[q] = a
  
  if verbose:
    print(form_data)
  user_agent = {'Referer':url,
                'User-Agent': "Mozilla/5.0 (X11; Linux i686) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/28.0.1500.52 Safari/537.36"}
  with requests.post(submit_url, data=form_data, headers=user_agent) as response:
    return response.status_code

async def get_questions(url):
  """Extrai questões e seus títulos de um formulário

  Args:
    url (string): URL do formulário

  Returns:
    dic: Relação de títulos e questões do formulário
  """
  res = urllib.request.urlopen(url)
  soupHtml = BeautifulSoup(res.read(), 'html.parser')
  get_labels = lambda f: [v for k, v in f.attrs.items() if 'label' in k]
  get_label = lambda f, n: get_labels(f)[0] if len(get_labels(f))>0 else n
  all_questions = soupHtml.form.findChildren(attrs={'name': lambda x: x and x.startswith('entry.')})
  all_div_labels = soupHtml.form.findChildren(attrs={'class': lambda x: x and x.startswith('freebirdFormviewerViewItemsItemItemTitle ')})
  all_labels = [str(k).strip() for k, v in all_div_labels]
  if len(all_labels) == len(all_questions):
    return {get_label(q, l): q['name'] for q, l in zip(all_questions, all_labels)}   
  else:
    raise Exception("A quantidade de questões e títulos não é a mesma")
    
  
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
      action="store_true",
      help="Retorna a relação de títulos e inputs das questões do formulário")
    groupExcl.add_argument(
      "--post",
      dest="op",
      action="store_false",
      help="Submete o formulário com as respostas")

    groupPost = parser.add_argument_group("POST")
   
    parser.add_argument(
      dest="url",
      help="URL do formulário",
      type=str,
      metavar="URL")
    
    groupPost.add_argument(
      "--questions",
      "-q",
      dest="questions",
      help="Questões do formulário no formato 'entry.99999[*]'",
      nargs="+",
      required='--post' in sys.argv,
      metavar="QUESTIONS")

    groupPost.add_argument(
      "--answers",
      "-a",
      dest="answers",
      help="Respostas do formulário",
      nargs="+",
      required='--post' in sys.argv,
      metavar="ANSWERS")

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
    if args.op:
      print("Questões:")
      print(asyncio.run(get_questions(args.url)))
    else:
      print("A submeter...")
      print("Resultado: {}".format(asyncio.run(post_answers(args.url, args.questions, args.answers))))
     
    _logger.info("Script ends here")


def run():
  """Entry point for console_scripts
  """
  main(sys.argv[1:])


if __name__ == "__main__":
  run()
