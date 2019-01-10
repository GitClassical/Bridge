# bridge-tools
Tools to and resources to facilitate data entry for The Bridge.

=======DATA=======
-morph_dataset_ORIGINAL.zip
-morph_dataset_SAMPLE.txt

=======CODE=======
sort_by_lemma.py
    Sorts the latin words in a Project Morpheus XML file, by lemma.

    USAGE:
    python sort_by_lemma.py XMLFILE

    ADDITIONAL INFO:
    This script is easily modified to sort by different word properties, by
      changing which element is used as a key for sorting.
    Uses the etree module of the standard python xml library to do XML parsing.
    (See the note on Project Morpheus XML files for details on XML structure.)

autoLemma.py
    Identifies possible lemmata and synopses for a set of latin words.

    Matches input words against a dataset of latin forms (e.g., that of Project
      Morpheus).  Can take as input either a spreadsheet or a txt file; for txt
      input, a new spreadsheet is generated showing the unput words and matched
      info about them.

    USAGE:
        python autoLemma.py [-a] TARGETFILE DATAFILE [--uniquesOnly=True/False]
    
    USAGE NOTES:
    TARGETFILE is a file containing the words to be lemmatized.
      Can be either a .xlsx spreadsheet, or a txt file.
      If .xlsx, the lemmata are written to a column in the file.
      If .txt, a new .xlsx file is created and the input words+lemmata
      are written in adjacent columns.
    DATAFILE is an xml file containing a dataset of latin words.
      This dataset must be structured like the Project Morpheus dataset:
      the root element, <analyses>, has a set of child <analysis> elements.
      Each <analysis> element corresponds to a unique latin word; that is,
      a unique combination of form (given by <form>), lemma (<lemma>),
      and synopsis (<pos>, <number>, etc).
    --uniquesOnly is an optional flag whether DATAFILE contains only unique
      latin forms.  Here we define "unique" as "having only one possible
      lemma", even if there are multiple possible synopses.
      If --uniquesOnly is passed, the program does not generate a list of
      ambiguous lemmata for the words in TARGETFILE.
    -a is an optional flag specifying whether to include ambiguous lemmata.
  

