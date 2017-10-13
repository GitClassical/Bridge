#!/usr/bin/env python
# coding=utf-8

"""
Identifies possible lemmata and snopyses for a set of Latin or Greek words.

Uses the lemmatizer provided by the Classical Languages Toolkit.
Must take a plain text file (.txt) as input; as output, a new spreadsheet
is generated showing the input words and matched info about them.

Locations should be separated from the rest of the text by square brackets. 
Examples: [1], [2.2], [3.1.2], ... If there are no location markers, then
line numbers will be used.

Note that docopt, openpyxl, regex, and the Classical Languages Toolkit must be 
installed as dependencies, as well as the relevant corpora: refer to
http://docs.cltk.org/en/latest/importing_corpora.html#importing-a-corpus.

Usage:
    autoLemma.py (greek|latin) [options] [--] <file> ...
    
Arguments:
    language                     only Greek and Latin are currently supported
    <file>                       one or more plain text files to be lemmatized

Options:
    -h --help                    show this help message and exit
    -a, --include-ambiguous      choose most likely lemma for an ambiguous form
    -f <format>, --l <format>    format of lemmata [default: cltk]       # TODO
    -o <file>, --output <file>   output file (uses input filename by default)
    --append                     append onto an existing output file
    --dir <path>                 prefixed onto file addresses for input/output
    -e, --echo                   print values of output spreadsheets to console
    --force-vi                   force lemmata to use "V, I" in place of "U, J"
    --force-ui                   force lemmata to use "U, I" in place of "V, J"
    --force-lowercase-lemmata    force lemmata to be lowercase
    --force-no-trailing-digits   force lemmata to not contain trailing digits
    --force-no-punctuation       force lemmata to not contain punctuation
    --force-uppercase-lemmata    force lemmata to be uppercase
    --group-by (section|location|file)  split-into groupings [default: section]
    --formulae                   also populate cells with formulae
    --split-into (files|sheets)  split output by (section) into sheets or files
    --text-name <title>          text name (for multiple input or output files)
    --output-sheet <name>        sheet name for output [default: LEMMATA MATCH]
    --use-line-numbers           include line numbers in word locations
    --use-sections               include column with simple section numbers
    --use-detailed-sections      include column with full location w/o line #s
"""

from sys import exit
from os.path import normpath, splitext, commonprefix, basename
from mimetypes import guess_type
from collections import namedtuple
from itertools import groupby
from unicodedata import normalize
from warnings import warn
from string import digits
import codecs

import utils.excel as excel

# https://pypi.python.org/pypi/regex
# this is to recognize accented Greek characters as alphabetic
import regex

# https://github.com/docopt/docopt
from docopt import docopt

# openpyxl.readthedocs.io
from openpyxl.utils import get_column_letter

# http://docs.cltk.org/en/latest/latin.html#lemmatization
from cltk.stem.lemma import LemmaReplacer 
from cltk.stem.latin.j_v import JVReplacer
from cltk.tokenize.word import WordTokenizer, nltk_tokenize_words

# TODO these should be patched in CLTK
IGNORED_LEMMATA = ['publica', 'tanto', 'multi', 'verro', 'medio', 'privo', 
                   'consento', 'quieto', 'mirabile', 'retineo', 'subeo', 
                   'Arruntius', 'disparo', 'prius', 'scelerato'] 

Word = namedtuple('Word', ['form', 'lemma', 'location'])
Word.__doc__ = \
    """
    A `Word` is a tuple with `form`, `lemma`, and `location` coordinates.

    Coordinates:
        form (str): the inflected word
        lemma (str or NoneType): the lemma which corresponds to `form`
        location (str): a string indicating the location of the word within
            a certain text; of the form 1, 2, 3, ... or 1.1, 1.2, 2.1, ...
    """
    
Location = namedtuple('Location', ['label', 'text'])
Location.__doc__ = \
    """
    A `Location` is a tuple with `label` and `text` coordinates.

    Coordinates:
        label (str): the label for the location; of the form 1, 2, 3, ...,
            or 1.1, 1.2, 2.1, ...
        text (str): the text contained in this location
    """
    
Column = excel.Column
Column.__doc__ += \
    """
    In this script, `find_value` coordinates of Column tuples will be passed
    Word tuples for their first parameter.
    """
    
CHECK_FORMULA = "=VLOOKUP({}{},'LEMMATA-Vocab'!A:A,1,FALSE)"
# COUNT_FORMULA = "=ROW({}{}) - 1"
DISPLAY_FORMULA = "=VLOOKUP({}{},'LEMMATA-Vocab'!C:E,2,FALSE)"
SHORTDEF_FORMULA = "=VLOOKUP({}{},'LEMMATA-Vocab'!C:E,3,FALSE)"
LONGDEF_FORMULA = "=VLOOKUP({}{},'LEMMATA-Vocab'!C:F,4,FALSE)"
LOCALDEF_FORMULA = "=VLOOKUP({}{},'LEMMATA-Vocab'!C:G,5,FALSE)"

FORMULAE_COLUMNS = [
    "CHECK", "DISPLAY LEMMA", "SHORTDEF", "LONGDEF", "LOCALDEF"
]

OUTPUT_COLUMNS = [
    Column("CHECK", 1, 
           lambda __, row: CHECK_FORMULA.format(LEMMA_COLUMN_LETTER, row)),
    Column("TITLE", 2, 
           lambda word, __: word.lemma if word.lemma is not None else ''),
    Column("TEXT", 3, 
           lambda word, __: word.form),
    Column("LOCATION", 4, 
           lambda word, __: word.location),
    Column("SECTION", 5, 
           lambda word, __: sectionFromWord(word)),
    Column("RUNNING COUNT", 6,
           lambda __, row: str(row)),
    Column("DISPLAY LEMMA", 7,
           lambda __, row: DISPLAY_FORMULA.format(LEMMA_COLUMN_LETTER, row)),
    Column("SHORTDEF", 8,
           lambda __, row: SHORTDEF_FORMULA.format(LEMMA_COLUMN_LETTER, row)),
    Column("LONGDEF", 9,
           lambda __, row: LONGDEF_FORMULA.format(LEMMA_COLUMN_LETTER, row)),
    Column("LOCALDEF", 10,
           lambda __, row: LOCALDEF_FORMULA.format(LEMMA_COLUMN_LETTER, row))
]

OUTPUT_COLUMNS_WITHOUT_FORMULAE = [
    column for column in OUTPUT_COLUMNS if column.name not in FORMULAE_COLUMNS
]

LEMMA_COLUMN_LETTER = get_column_letter(
    excel.getColumnByName(OUTPUT_COLUMNS, "TITLE").number
)

class NLTKTokenizer(WordTokenizer):
    """
    A wrapper for the `nltk_tokenize_words` function that inherits from the
    cltk.tokenize.word.WordTokenizer class. This could be expanded to implement
    some special tokenization for Greek if needed.
    """
    def __init__(self, language):
        self.language = language
    def tokenize(self, string):
        return nltk_tokenize_words(string)
      
def processUnicodeDecomposition(string, *functions):
    """
    Apply all functions in arguments to the canonicial decomposition of the 
    unicode string `string` (i.e. with combining characters such as accents
    separated from the characters they combine with).
    
    Parameters:
        string (str): a Unicode string
        *functions (List[Callable]): variable length list of functions
        
    Returns:
        (str): the canonical normalization of `string` after each function in 
            `*functions` has been applied to its decomposition
    """
    decomposition = normalize('NFD', string)
    for function in functions:
        decomposition = function(decomposition)
    return normalize('NFC', decomposition)

def removeMacrons(string):
    """"
    Removes the macrons from any macron-ed characters in a string.
    
    Parameters:
        string (str): the string whose macrons are to be removed
            macrons must be separated as combining characters
        
    Returns:
        (str): `string` without any macrons
    """
    return regex.sub('\u0304', '', string)
  
def removeDiareses(string):
    """"
    Removes any diareses in string.
    
    Parameters:
        string (str): the string whose diareses are to be removed
            diareses must be separated as combining characters
        
    Returns:
        (str): `string` without any diareses
    """
    return regex.sub('\u0308', '', string)
  
def changeGraveAccents(string):
    """
    Changes any grave accents in `string` to acute accents.
    
    Parameters:
        string (str): the string whose grave accents are to be replaced
            grave accents must be separated as combining characters
        
    Returns:
        (str): `string` with acute accents in place of graves
    """
    return regex.sub('\u0300', '\u0301', string)

def sectionFromWord(word):
    """
    Returns the section associated with the location coordinate of the Word
    tuple `word`. The section is given by the first identifiable group of 
    characters in a location split by delimeters or mixed between alphabetic
    and numeric characters.
    
    For example, the locations '1.1' and '1.2' are both considered to belong to
    section '1'; likewise, '9a' and '9b' both belong to section '9'.
    
    Parameters:
        word (Word): the Word tuple whose section is to be found
    
    Returns:
        section (str): the section associated with the location of `word`
    """
    match = regex.match(r'(?i)[0-9]+|[A-Z]+|[Α-Ω]+',word.location)
    return match.group() if match is not None else None

def detailedSectionFromWord(word):
    """
    Returns the detailed section associated with the location coordinate of 
    the Word tuple `word`. The detailed section is given by all but the last 
    identifiable groups of characters in a location split by delimeters or 
    mixed between alphabetic and numeric characters.
    
    For example, the locations '1.1.3' and '1.2.4' are both considered to 
    belong to sections '1.1' and '1.2' respecively; likewise, '9a1' and '9b1' 
    belong to sections '9a' and '9b'.
    
    Parameters:
        word (Word): the Word tuple whose section is to be found
    
    Returns:
        section (str): the detailed section associated with the location 
            coordinate of `word`
    """
    match = regex.match(r'(?i)[0-9]+|[A-Z]+|[Α-Ω]+', word.location[::-1])
    if match is not None:
        return word.location[:-1 * match.end()].rstrip('.')
    
def groupWordsBySection(words, detailed=False):
    """
    Groups the Word tuples of `words` by the sections associated with their
    location coordinates as identified by `sectionFromWord` or by
    `detailedSectionFromWord`.
    
    Parameters:
        words (List[Word]): an iterable over Word tuples sorted by location
        detailed (bool): whether to use the `detailedSectionFromWord` function
    
    Returns:
        an iterator over tuples of section (str): subiterable over `words`
    """
    return groupby(words, sectionFromWord)
  
def groupWordsByLocation(words):
    """
    Groups the Word tuples of `words` by unique location coordinates.
    
    Parameters:
        words (List[Word]): an iterable over Word tuples sorted by location
    
    Returns:
        an iterator over tuples of location (str): subiterable over `words`
    """
    return groupby(words, lambda word: word.location)
  
def lemmatizeToken(token, lemmatizer):
    """
    Returns the lemma corresponding to the token `token`.
    
    `Token` should be already formatted for CLTK lemmatization (i.e.
    tokenized into a single word, with JV-Replacement if Latin,
    in Unicode format for Greek).
    
    Parameters:
        string (string): the string being lemmatized
        lemmatizer (cltk.stem.lemma.LemmaReplacer): the lemmatizer to use to
            identify the lemma of `token`
        
    Returns:
        lemma (string): the lemma which corresponds to `token`,
            or `None` if no such lemma was found
        
    Raises:
        ValueError if `token` contains non-alphabetic characters, 
            multiple lemmata, or only whitespace
    """
    token = token.lower()
    lemmata = lemmatizer.lemmatize(token, default = '')
    if len(lemmata) > 1:
        raise ValueError("'{}' contains multiple lemmata".format(token))
    elif len(lemmata) == 0:
        raise ValueError("'{}' is empty or nonalphabetic.".format(token))

    lemma = lemmata[0]
    if not lemma:
        # To prevent problems with case (which should be fixed in CLTK with the lemmatizer rewrite)
        lemma = lemmatizer.lemmatize(token.capitalize(), default = '')[0]
    if lemma in IGNORED_LEMMATA:
        lemma = ""         # For some known bugs with CLTK
    
    return lemma if lemma else None
            
def locationsFromFile(file, *, use_line_numbers = False):
    """
    Iterates over unique locations as indicated by section delimiters or line
    numbers in the text file `file`.
    
    Parameters:
        file (io.TextIOWrapper): the file being read
        use_line_numbers (bool): whether to append line numbers to locations
            when section delimiters are present
        
    Yields:
        location (Location): a Location tuple of the next unique location in 
            `file` and the text it contains
    """
    line_number = 1
    text, section = '', ''
    
    def getFormattedLabel():
        if section and use_line_numbers:
            return '{}.{}'.format(section, line_number)
        else:
            return section if section else str(line_number)
    
    for line_text in file:
        # split by lines
        if regex.search(r'[0-9]+$', line_text):
            # found inline line number
            line_number = int(regex.search(r'[0-9]+$', line_text).group())
        for string in regex.split(r'(\[[0-9.]+\])', line_text):
            # split by delimeters
            # http://stackoverflow.com/questions/2136556/in-python-how-do-i-split-a-string-and-keep-the-separators
            if regex.search(r'\[[0-9.]+\]', string):
                # found section delimiter
                if text and section and not text.isspace():
                    yield Location(getFormattedLabel(), text)
                # write section delimeter to location
                text, section, line_number = '', string[1:-1], 1
            elif string and not string.isspace():
                text += string
        if text and (use_line_numbers or not section):
            yield Location(getFormattedLabel(), text)
            text, line_number = '', line_number + 1
            
    yield Location(getFormattedLabel(), text)
  
def wordsFromFile(file, lemmatizer, *, use_line_numbers = False):
    """
    Extracts words as Word tuples from the text file `file`.
    
    Parameters:
        file (io.TextIOWrapper): the file being read
        lemmatizer (cltk.stem.lemma.LemmaReplacer): the lemmatizer to use to
            identify lemmata of words in `file`
        
    Yields:
        word (Word): a Word tuple of the next form to appear in `file`
    """
    jv_replacer = JVReplacer()
    try:
        tokenizer = WordTokenizer(lemmatizer.language)
    except AssertionError:
        # Language of `lemmatizer` does not support CLTK tokenization
        # Clitics must be separated by spacing
        tokenizer = NLTKTokenizer(lemmatizer.language)
        
    locations = locationsFromFile(file, use_line_numbers=use_line_numbers)
    for location, text in locations:
        if lemmatizer.language == 'latin':
            text = jv_replacer.replace(text)
            text = processUnicodeDecomposition(text, removeMacrons)
        elif lemmatizer.language == 'greek':
            text = processUnicodeDecomposition(text, removeDiareses, 
                                               changeGraveAccents)
        for token in tokenizer.tokenize(text):
            for form in regex.split(r'(?:\P{L}+)', token):
                # split around punctuation
                try:
                    lemma = lemmatizeToken(form, lemmatizer)
                    yield Word(form, lemma, location)
                except ValueError:
                    # token contains non-alphabetic characters which
                    # aren't leading or trailing, or contains no
                    # alphabetic characters
                    continue

def wordsFromPathList(paths, lemmatizer, **kwargs):
    """
    Extracts words as Word tuples from the files given by `paths`. Note: 
    removes paths from `paths` successively during iteration.
    
    Parameters:
        paths (List[str]): each path must be the address of a plain text file
        lemmatizer (cltk.stem.lemma.LemmaReplacer): the lemmatizer to use to
            identify lemmata of words in `file`
        
    Yields:
        word (Word): the next Word tuple to appear in the files in `paths`
    """
    while len(paths) > 0:
        path = paths.pop(0)
        file_type = guess_type(path)[0]
        if file_type != 'text/plain':
            warn("Only plain text input is supported: "
                 "{} appears to be {}.".format(path, file_type))
        print("Loading {}".format(path))
        try:
            with codecs.open(path, 'r', 'utf-8') as file:
                words = wordsFromFile(file, lemmatizer, **kwargs)
                yield from words
        except IOError:
            exit("Could not find file in path {}".format(path))
                    
def autoLemma(args, *, lemmatizer=None, wordsFromPathList=wordsFromPathList):
    """
    Generates lemmatized spreadsheets from command-line arguments given by 
    `args` (see docstring of `autoLemma.py`) using the function
    `wordsFromPathList` to extract Word tuples from the paths in 
    `args['<file>']`.
    
    Parameters:
        args (dict of str:str): see docstring of autoLemma.py
        lemmatizer (cltk.stem.lemma.LemmaReplacer): the lemmatizer to use to
            identify lemmata of words
        wordsFromPathList (Callable): returns an iterable over Word tuples from 
            file paths
    """
    if lemmatizer is None:
        lemmatizer = LemmaReplacer('latin' if args['latin'] else 'greek', 
                                   include_ambiguous=args['--include-ambiguous'])
    # Find name of text
    if args['--output']:
        text_name, text_extension = splitext(args['--output'])
        if text_extension != 'xlsx':
            warn("autoLemma.py only outputs .xlsx files")
    elif args['--text-name']:
        text_name = args['--text-name']
    else:
        text_name = splitext(basename(commonprefix(args['<file>'])))[0]
        if not text_name: text_name = 'output'
        
    # Create input/output path
    path_prefix = (normpath(args['--dir']) + '/') if args['--dir'] else ''
    
    # Extract Word tuple iterators from input
    paths = [path_prefix+path for path in args['<file>']]
    try:
        words = wordsFromPathList(paths, 
                                  lemmatizer=lemmatizer,
                                  use_line_numbers=args['--use-line-numbers'])
    except IOError:
        exit("Error opening one or more files! {!s}".format(args['file']))
        
    # Group words if necessary by specified means
    if args['--split-into']:
        if args['--group-by'] == 'file':
            groups = groupby(words, lambda __: len(args['<file>']) - len(paths))
        elif args['--split-into'] and args['--group-by'] == 'location':
            groups = groupWordsByLocation(words)
        elif args['--split-into'] and args['--group-by'] == 'section':
            if args['--use-detailed-sections']:
                groups = groupWordsBySection(words, detailed=True)
            else:
                groups = groupWordsBySection(words)
        else:
            exit("Invalid --group-by argument")
          
    # Write words and matched lemmata to new spreadsheets
    if args['--formulae']:
        columns = OUTPUT_COLUMNS
    else:
        columns = OUTPUT_COLUMNS_WITHOUT_FORMULAE 
        
    if args['--use-detailed-sections']:
        excel.replaceColumnFunction(columns, "SECTION",
                                    lambda word, __: detailedSectionFromWord(word))
    elif not args['--use-sections']:
        columns = [ column for column in columns if column.name != "SECTION" ]
        
    if args['--force-lowercase-lemmata']:
        excel.wrapColumnFunction(columns, "TITLE", lambda lemma: lemma.lower())
        
    if args['--force-uppercase-lemmata']:
        excel.wrapColumnFunction(columns, "TITLE", lambda lemma: lemma.upper())
        
    if args['--force-no-trailing-digits']:
        excel.wrapColumnFunction(columns, "TITLE",
                                 lambda lemma: lemma.rstrip(digits))
        
    if args['--force-no-trailing-digits']:
        excel.wrapColumnFunction(columns, "TITLE",
                                 lambda lemma: regex.sub(r'\P{L}', '', lemma))
        
    if args['--force-vi'] or args['--force-ui']:
        if args['--force-vi']:
            replacements = [('u', 'v'), ('U', 'V'), ('j', 'i'), ('J', 'I')]
        elif args['--force-ui']:
            replacements = [('v', 'u'), ('V', 'U'), ('j', 'i'), ('J', 'I')]
        def replace(lemma):
            for pattern, repl in replacements:
                lemma = regex.sub(pattern, repl, lemma)
            return lemma
        excel.wrapColumnFunction(columns, "TITLE", lambda lemma: replace(lemma))
        
    if args['--split-into'] == 'files':
        excel.saveGroupsToSpreadsheets(
            groups, columns, path=path_prefix, file_prefix=text_name, 
            append=args['--append'], sheet_title=args['--output-sheet'],
            echo=args['--echo']
        )
    elif args['--split-into'] == 'sheets':
        excel.saveGroupsToSpreadsheet(
            groups, columns, text_name, path=path_prefix, echo=args['--echo'], 
            append=args['--append'], sheet_prefix=args['--output-sheet']
        )
    else:
        excel.saveDataToSpreadsheet(
            words, columns, text_name, path=path_prefix, echo=args['--echo'], 
            append=args['--append'], sheet_title=args['--output-sheet'],
        )

if __name__ == '__main__':
    autoLemma(docopt(__doc__))
    exit()