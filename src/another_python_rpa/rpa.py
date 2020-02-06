# -*- coding: utf-8 -*-
"""
This is a skeleton file that can serve as a starting point for a Python
console script. To run this script uncomment the following lines in the
[options.entry_points] section in setup.cfg:

    console_scripts =
         rpa = another_python_rpa.rpa:run

Then run `python setup.py install` which will install the command `rpa`
inside your current environment.
Besides console scripts, the header (i.e. until _logger...) of this file can
also be used as template for Python modules.

Note: This skeleton file can be safely removed if not needed!
"""

import argparse
import sys
import logging
import asyncio
import another_python_rpa

from another_python_rpa import __version__, xlsxutils, formutils

__author__ = "Arthur Aleksandro Alves Silva"
__copyright__ = "Arthur Aleksandro Alves Silva"
__license__ = "mit"

_logger = logging.getLogger(__name__)


async def rpa(url, file, name=None, default_name ="Arthur Aleksandro Alves Silva",):
  """RPA function
    start and control view
  """
  try:
    fileRel = xlsxutils.extract(file)
    questsRel = await formutils.get_questions(url)
    finalName = name if name else default_name
    tasks = []
    for row in fileRel['rows']:
      row["Nome completo do candidato"] = finalName
      quests = []
      ans = []
      for c, v in row.items():
        if c in questsRel:
          quests.append(questsRel[c])
          ans.append(v)
      tasks.append(formutils.post_answers(url, quests, ans))
    result = []
    while len(tasks):
      done, tasks = await asyncio.wait(tasks, return_when=asyncio.FIRST_COMPLETED)
      for task in done:
        result.append(task.result())
    return result
  except:
    # todo: tratar melhor depois
    raise Exception("Algo de errado aconteceu")

  # print([[(questsRel[c], v) for c, v in k.items() if c in questsRel] for k in fileRel['rows']])
""" 
  for f in fileRel:
     name if name else default_name """


def parse_args(args):
  """Parse command line parameters

  Args:
    args ([str]): command line parameters as list of strings

  Returns:
    :obj:`argparse.Namespace`: command line parameters namespace
  """
  parser = argparse.ArgumentParser(
    description="RPA demonstration")
  parser.add_argument(
    "--name",
    "-n",
    dest="name",
    help="Nome do candidato",
    type=str,
    metavar="NOME")
  parser.add_argument(
    dest="url",
    help="URL do formulário",
    type=str,
    metavar="URL")
  parser.add_argument(
    "--file",
    "-f",
    dest="file",
    help="Arquivo de planilha",
    type=str,
    metavar="FILE")
  parser.add_argument(
    "--version",
    action="version",
    version="another-python-rpa {ver}".format(ver=__version__))
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
  _logger.debug("Starting view and controllers...")
  # print("Argumentos recebidos: {}".format(args))
  file = args.file
  if not file:
    print("Informe o arquivo (caminho completo): ")
    file = input()
  loop = asyncio.get_event_loop()
  try:
    loop.run_until_complete(rpa(args.url,file, args.name))
  finally:
    loop.close()
  _logger.info("Script ends here")


def run():
    """Entry point for console_scripts
    """
    main(sys.argv[1:])


if __name__ == "__main__":
    run()
