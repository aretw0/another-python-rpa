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
from concurrent.futures import ThreadPoolExecutor

from tkinter import *
from tkinter.ttk import Progressbar
from tkinter import ttk, filedialog, messagebox

from another_python_rpa import __version__, xlsxutils, formutils

__author__ = "Arthur Aleksandro Alves Silva"
__copyright__ = "Arthur Aleksandro Alves Silva"
__license__ = "mit"

_logger = logging.getLogger(__name__)


async def rpa(window, url, file, name, bar=None):
  """RPA function
  Args:
    window (tkinter): window para atualizar
    url (string): URL do formulário
    file (string): Arquivo da planilha
    name (string): Nome do candidato
    bar (widget): Barra de progresso

  Returns:
    int: Códido HTTP da submissão
  """
  fileRel = xlsxutils.extract(file)
  questsRel = await formutils.get_questions(url)
  tasks = []
  with ThreadPoolExecutor(max_workers=10) as executor:
    loop = asyncio.get_event_loop()
    for row in fileRel['rows']:
      row["Nome completo do candidato"] = name
      quests = []
      ans = []
      for c, v in row.items():
        if c in questsRel:
          quests.append(questsRel[c])
          ans.append(v)
      tasks.append(
        loop.run_in_executor(
          executor,
          formutils.post_answers,
          *(url, quests, ans)
        )
      )
    result = {
      "ok": 0,
      "notOk": 0
    }
    lenTasks = len(tasks)
    if bar:
      bar.configure(length=lenTasks)
      bar['value'] = 0
    for response in await asyncio.gather(*tasks):
      result["ok" if response == 200 else "notOk"] += 1
    return result
    """ 
    while len(tasks):
      for asyncio.gather(*tasks)
      done, tasks = await asyncio.wait(tasks, return_when=asyncio.FIRST_COMPLETED)
      for task in done:
        print("Done!")
        bar['value'] += 1
        window.update()
        result["ok" if task.result() == 200 else "notOk"] += 1
    return result
    """

window = None
nameEntry = None
fileEntry = None
formEntry = None

nameVar = None
fileVar = None
formVar = None

submitBtn = None
fileBtn = None

submitBar = None

def rpa_window(url,file,name="", default_name ="Arthur Aleksandro Alves Silva"):
  # fonte e texto de apresentação
  useFont = "Segoe UI"
  presentText = "Esta aplicação automatiza a submissão em massa de formulários online (Google Forms) usando dados de uma planilha."
  widthEntries = 27

  window = Tk()
  window.resizable(0,0)
  window.title("Another Python RPA")
  # window.geometry('800x600')
  
  # Labels
  Label(window, text="Bem vindo!", font=(useFont, 15)).grid(row=0, column=0, columnspan=2,
                                                          sticky=W+E+N+S, padx=5, pady=(5,0))  
  Label(window, text=presentText, font=(useFont, 13), wraplength=380, justify='center').grid(
    row=1, column=0, columnspan=2, sticky=W+E+N+S, padx=5, pady=(0,20))
  
  nameVar = StringVar()
  fileVar = StringVar()
  formVar = StringVar()

  nameVar.set(name if name else default_name)
  fileVar.set(file if file else "")
  formVar.set(url if url else "")


  nameEntry = Entry(window, textvariable=nameVar, font=(useFont, 12))
  formEntry = Entry(window, textvariable=formVar, font=(useFont, 12))
  fileEntry = Entry(window, textvariable=fileVar, font=(useFont, 12))

  Label(window, text="Nome do candidato", font=(useFont, 12)).grid(columnspan=2,row=2, sticky=W, padx=10)
  nameEntry.grid(column=0, columnspan=2, padx=10, sticky=W+E)

  Label(window, text="Formulário (URL)", font=(useFont, 12)).grid(columnspan=2,row=4, sticky=W, padx=10, pady=(5,0))
  formEntry.grid(column=0, columnspan=2, padx=10, sticky=W+E)

  Label(window, text="Planilha", font=(useFont, 12)).grid(row=6, sticky=W, padx=(10,0), pady=(5,0))
  fileEntry.grid(column=0, padx=(10,0), sticky=W+E, pady=(0,15))
  
  def getFile():
    fileName = filedialog.askopenfilename(filetypes = (("Excel files","*.xlsx"),("all files","*.*")))
    fileVar.set(fileName if fileName else fileVar.get())
  
  fileBtn = Button(window, text="Selecionar", command=getFile, font=(useFont, 0))
  fileBtn.grid(column=1, row=7, sticky=W+E, pady=(0,15), padx=(5,10))

  def submit():
    if fileVar.get() and formVar.get() and nameVar.get():
      nameEntry.configure(state='disabled')
      formEntry.configure(state='disabled')
      fileEntry.configure(state='disabled')
      fileBtn.configure(state='disabled')
      submitBtn.configure(state='disabled')
      loop = asyncio.get_event_loop()
      try:
        messagebox.showwarning('ATENÇÃO', 'É bem provável que esta aplicação fique sem responder durante o processamento. Peço que apenas aguarde e se tiver controle sobre o fomulário verifique se as respostas estão incrementando.')
        future = asyncio.ensure_future(rpa(window, formVar.get(), fileVar.get(), nameVar.get(), submitBar))
        result = loop.run_until_complete(future)
        messagebox.showinfo('Finalizado', '{} formulários submetidos com sucesso, {} falharam'.format(result['ok'],result['notOk']))
      except Exception as err:
        messagebox.showerror('Erro', 'Algo errado aconteceu :(')
      finally:
        loop.close()
        nameEntry.configure(state='normal')
        formEntry.configure(state='normal')
        fileEntry.configure(state='normal')
        fileBtn.configure(state='normal')
        submitBtn.configure(state='normal')
    else:
      messagebox.showinfo('Campo(s) Inválido(s)', 'Preencha tudo corretamente e tente novamente')


  submitBtn = Button(window, text="Submeter", command=submit, font=(useFont, 12))
  submitBtn.grid(columnspan=2, padx=10, sticky=W+E)

  style = ttk.Style()
  style.theme_use('default')
  style.configure("blue.Horizontal.TProgressbar", background='blue')

  submitBar = Progressbar(window, style='blue.Horizontal.TProgressbar' )
  submitBar.grid(padx=10, pady=10, columnspan=2, sticky=W+E) 

  window.mainloop()


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
    "--url",
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
  rpa_window(args.url,args.file,args.name)
  """ 
  file = args.file
  if not file:
    print("Informe o arquivo (caminho completo): ")
    file = input()
  """
  """
  loop = asyncio.get_event_loop()
  try:
    loop.run_until_complete(rpa(args.url,file, args.name))
  finally:
    loop.close() """
  _logger.info("Script ends here")


def run():
    """Entry point for console_scripts
    """
    main(sys.argv[1:])


if __name__ == "__main__":
    run()
