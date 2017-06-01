# -*- coding: utf-8 -*-
"""
This code exists for those who want to take information from an Excel spreadsheet, such
as survey data, and compile it into a document. This is Python code which outputs a .tex
file to be compiled with LaTeX. See https://github.com/gknave/ExcelToLaTeX for instructions.

Created on Tue Feb 14 21:49:14 2017

@author: Gary
"""


from pandas import read_excel, read_csv

# This is a dictionary of unicode characters for conversion to LaTeX format.
ut8dict = {'\u03bc': '\\(\\mu\\)',
     '\u03b2': '\\(\\beta\\)',
     '\u03b3': '\\(\\gamma\\)',
     '\u03b1': '\\(\\alpha\\)',
     '\u2264': '\\(\\leq\\)',
     '\u03b4': '\\(\\delta\\)'}

def charFormat(string):
  """Formats a string for LaTeX use.
  
  Called by substitute(). Uses ut8dict above to convert unicode characters to their LaTeX format.
  Also converts %&@ to their LaTeX text equivalent: \% \& and \@. Does not
  convert $ to \$ such that equations may be included in the input.

  Note that %, &, and @ are converted to their text equivalents, so problems
  may arise if they are used as their TeX commands.
  
  Parameters
  ----------
  string : string
    A string of text to be formatted for LaTeX output.

  Returns
  -------
  out : string
    A string with unicode and special LaTeX characters converted to text

  Examples
  --------
  >> charFormat('Romeo & Juliet')
  'Romeo \& Juliet'
  >> charFormat('100% effective!')
  '100\% effective!'
  """
  out = ''
  if not type(string) == str:
    return out
  for a in string:
    if a in '%&@':
      out += '\\' + a
    elif a in ut8dict.keys():
      out += ut8dict[a]
    else:
      out += a
  return out

def substitute(string, df, keys = [], input='', keycheck=False):
  """Function for substituting pandas dataframe content into a LaTeX document.

  Called by toLatex(). Scans a string for text of the form {{key}} and replaces it with
  df[key]. The dataframe is assumed to be a single row, representing
  one row of entries from an Excel spreadsheet.

  Parameters
  ----------
  string : a string
    This is the raw string to be processed. Any text in double brackets {{}}
    is assumed to be in df.keys().

  df : a one row pandas dataframe
    Contains the Excel content to be inserted into string anywhere double
    brackets appear.

  keys : list of strings, optional, default: []
    This is used for checking the keys that substitute() found. Returned if 
    keycheck==True. Used for recursion.
  
  input : string, optional, default: ''
    The conversion of string is appended to input. Used for recursion.
  
  keycheck : [True, False], default: False
    Chooses whether keys are returned. Used for troubleshooting.

  Returns
  -------
  out : string
    LaTeX readable string containing substitutions from the input dataframe
    anywhere the input string contains double brackets {{key}}.
  
  keys : list of strings, optional
    If keycheck==True, the command outputs the keys found during substitution.
    Used for troubleshooting key errors.

  """
  first = False
  k = 0
  for a in string:
    if a == '{':
      if first:
        k += 1
        key = ''
        for b in string[k:]:
          if b == '}':
            k += 2
            break
          else:
            key += b
            k += 1
        keys.append(key)
        out += charFormat(df[key])
        return substitute(string[k:], df, keys=keys, input=out, keycheck=keycheck)
      else:
        k += 1
        first = True
    elif first:
      out += '{'+a
      k += 1
      first = False
    else:
      out += a
      k += 1
  if keycheck:
    return out, keys
  else:
    return out
      
def toLatex(sheet='sheet.xlsx', preamble='preamble.tex', entryStyle='style.tex', output='output.tex', sheetType='excel', endFile=None):
  """Creates a LaTeX document from an Excel spreadsheet.

  From an Excel spreadsheet, sheet, and two LaTeX files, preamble and entryStyle, creates 
  a LaTeX file, output. Substitutes values from sheet into entryStyle each time a double
  bracket {{key}} is found, where 'key' is a header in sheet.

  Parameters
  ----------
  sheet : string referencing .xlsx file, default: 'sheet.xlsx'
    If sheetType='csv', this may be a .csv file

  preamble : string referencing .tex file, default: 'preamble.tex'
    .tex file generating the preamble to the output .tex file.

  entryStyle : string referencing .tex file, default: 'style.tex'
    .tex file iteratively used for each row of sheet. Double brackets 
    {{key}} indicate an entry from a column in sheet labeled 'key'.

  output : string giving name of output .tex file, default: 'output.tex'
    This provides the name for the generated .tex file

  sheetType : ['excel', 'csv'], default: 'excel'
    Indicates whether sheet is a CSV or Excel document.

  endFile : string referencing .tex file or None, default: None
    Optional .tex file to be added to end of output
  
  Returns
  -------
  Creates .tex file named output.

  """
  if sheetType == 'excel':
    df = pd.read_excel(sheet)
  elif sheetType == 'csv':
    df = pd.read_csv(sheet)
  else:
    raiseError('sheetType must be "excel" or "csv"')

  prefile = open(preamble, 'r')
  pre = prefile.read()
  prefile.close()

  styleFile = open(entryStyle, 'r')
  style = styleFile.read()
  styleFile.close()
  
  if not endFile == None:
    endfile = open(endFile, 'r')
    end = endfile.read()
    endfile.close()

  file = open(output, 'w')
  file.write(pre)
  if '\\begin{document}' not in pre:
    file.write('\\begin{document}\n')
  for index in range(len(df)):
    temp = df.iloc[index]
    line = substitute(style, temp, [], '')
    file.write('\n' + line + '\n')
  
  if endFile == None:
    file.write('\\end{document}')
  elif '\\end{document}' in end:
    file.write('\n' + end + '\n')
  else:
    file.write('\n' + end + '\n\\end{document}')
  file.close()
