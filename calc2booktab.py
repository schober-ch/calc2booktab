#########################################################################
#                                                                       #
#                            Calc2BOOKTAB                               #
#                                                                       #
#                                                                       #
#       COPYRIGHT (c) 2012-2014 Christoph Schober                       #
#       (christoph.schober@ch.tum.de)                                   #
#                                                                       #
# This program is free software: you can redistribute it and/or modify  #
# it under the terms of the GNU General Public License as published by  #
# the Free Software Foundation, either version 3 of the License, or     #
# (at your option) any later version.                                   #
#                                                                       #
# This program is distributed in the hope that it will be useful,       #
# but WITHOUT ANY WARRANTY; without even the implied warranty of        #
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the         #
# GNU General Public License for more details.                          #
# You should have received a copy of the GNU General Public License     #
# along with this program.  If not, see <http://www.gnu.org/licenses/>. #
#                                                                       #
#########################################################################

# This script will convert the selected table in OpenOffice / LibreOffice
# to LateX-Code

# Import python modules
import uno
import csv
#import unohelper
import tempfile
import os
import re
from com.sun.star.awt import Selection


######################CUSTOM SETTINGS FOR THE SCRIPT######################
##########################################################################
##########################################################################
#change label, caption, etc for special scripts, e.g. vimlatex           #
#
label = "LABEL"                                                          #
caption = "CAPTION"                                                      #
placement = "ht!"                                                        #
#
#some other custom settings for the script:                              #
#
#Format first line as \multicolumn{1}{c} (change to "False" if unwanted!)#
multicol = True                                                          #
str_mult_col = " \\multicolumn{1}{c}{"                                     #
#
#Active column-definition with dcolumn-package                           #
#Column definition for dcolumn (see manual for more information!)        #
dcol_def = "D{.}{.}{-1}"                                                  #
#
#Fixed width for table ( \begin{tabular*}-enviroment                     #
#Define your width in .% of textwidth                                    #
str_fixed_width = "0.9"                                                    #
##########################################################################
#Definitions for LaTeX-Table Header and Footer                           #

str_begin_table = "\\begin{table}"
str_centering = "\\centering"
str_begin_tabular = "\\begin{tabular}"

str_end_tabular = "\\end{tabular}"
str_end_table = "\\end{table}"

str_begin_tabular2 = "\\begin{tabular*}"
str_end_tabular2 = "\\end{tabular*}"

##########################################################################
# Do NOT change anything below this line if you don't know what it means #
##########################################################################

# LATEX table header + footer: Search for "header" / "footer" in the code

# definition of some global variables

ctx = uno.getComponentContext()
smgr = ctx.ServiceManager

# function to check if string is a decimal
# Somewhat extented checking for typical scientific tables:
# True if:
#       ... s = 5.342
#       ... s = 5,342
#       ... s = 5,342(5)
#       ... s = 5,342 (5)
#       ... s = 5,342 this value sucks


def is_number(s):
    """Function to check if string is a decimal.
     Somewhat extented checking for typical scientific tables:
     True if:
           ... s = 5.342
           ... s = 5,342
           ... s = 5,342(5)
           ... s = 5,342 (5)
           ... s = 5,342 some text"""

    if s != "":
        if "," in s:
            s = s.replace(",", ".", 1)
        else:
            pass
        sn = re.split('[*()-+/\! ]', s)
        s = sn[0]
        try:
            float(s)
            return True
        except ValueError:
            return False
    else:
        return False


def calc2booktab_dcolumn(*vartuple):
    """calc2booktab_dcolumn adds functionality for numeric columns.

    Read the dcolumn-package manual for more information on what it
    does and how it is used (http://www.ctan.org/pkg/dcolumn).
    Please be aware that the script can only handle tables with one
    or zero columns of text. More text-columns will still be converted,
    but you have to adjust the dcolumn-definition in the code manually."""

    global dcolumn
    global fixed_width
    dcolumn = True
    fixed_width = False
    convert_code()


def calc2booktab_fixed(*vartuple):
    """ calc2booktab_fixed will format the table header with the
    tabular*-enviroment.

    The result will be a table with a fixed width.
    Default is 0.9\\textwidth. If you want to use another standard-value,
    you can change this at the beginning of the calc2booktab.py-file.
    Have a look at the LaTeX-manual if you don't know this enviroment yet
    and want to learn more about it.
    (for example: http://en.wikibooks.org/wiki/LaTeX/Tables)"""

    global dcolumn
    global fixed_width
    fixed_width = True
    dcolumn = False
    convert_code()


def calc2booktab_fixed_dcolumn(*vartuple):
    """ calc2booktab_fixed_dcolumn combines dcolumn with fixed_table."""

    global dcolumn
    global fixed_width
    fixed_width = True
    dcolumn = True
    convert_code()


def calc2booktab_basic(*vartuple):
    """ calc2booktab_basic does a standard conversion of the selected cells
    to valid latex code. Cells with numbers will be converted to math mode, e.g.
    cell a cell with "3.13" will be "$3.13$"."""

    global dcolumn
    global fixed_width
    dcolumn = False
    fixed_width = False
    convert_code()


def convert_code():
    """main function for handling calc-sheet and extracting the data"""

# global variable definitions for all functions, needed for multiple runs
# of the script during one office-section
    # align_final = ""
    global lines_final
    global header
    global footer
    global output
    global final_code
    global final_code_t
    global align_final
    align_final = ""
    data_list = []
    final_code = []
    final_code_t = ""

# Define temp-file
# using to dump csv table for easier formatting(?)

    output_file = tempfile.NamedTemporaryFile(delete=False)
    output = output_file.name
    output_file.close()

    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()
    o_selection = doc.getCurrentSelection()
    o_area = o_selection.getRangeAddress()
    frow = o_area.StartRow
    lrow = o_area.EndRow
    fcol = o_area.StartColumn
    lcol = o_area.EndColumn

# loop over selection
    new_data = []
    c_range = range(fcol, lcol + 1)
    r_range = range(frow, lrow + 1)
    nrcols = lcol - fcol

# single row loop for text alignment (if dcolumn = False!)
# loop will check the 2nd row of the selection and assign all user-made
# alignments, then check for standard-alignments (no specific selection by
# user). If standard found, cells with numbers will be "center" (c), cells
# with text will be "left" if firstrow, "center" if any other.

# check 1st column for num or text. If 1st col = num, expect all other
# cols to be numeric too. Sets dcolumn if true, else cell adjustment (see
# above). Has also problems with 2 top-lines (e.g. 1st + 2nd line with
# captions). Could be solved if script checks how many cols in first row
# are empty

    first_cell_text = False
    for i in c_range:
        if i == fcol:
            pos_cell = sheet.getCellByPosition(i, frow + 1)
            if is_number(pos_cell.String):
                if dcolumn:
                    align_final = "{" + str(nrcols + 1) + "}{" + dcol_def + "}"
                    break
                else:
                    alignment_c = pos_cell.HoriJustify
                    if "LEFT" in str(alignment_c):
                        align_final = align_final + "l"
                    elif "CENTER" in str(alignment_c):
                        align_final = align_final + "c"
                    elif "RIGHT" in str(alignment_c):
                        align_final = align_final + "r"
                    elif "STANDARD" in str(alignment_c):
                        align_final = align_final + "c"
            else:
                align_final = align_final + "l"
                first_cell_text = True
        else:
            pos_cell = sheet.getCellByPosition(i, frow + 1)
            if not dcolumn:
                alignment_c = pos_cell.HoriJustify
                if "LEFT" in str(alignment_c):
                    align_final = align_final + "l"
                elif "CENTER" in str(alignment_c):
                    align_final = align_final + "c"
                elif "RIGHT" in str(alignment_c):
                    align_final = align_final + "r"
                elif "STANDARD" in str(alignment_c):
                    align_final = align_final + "c"
            else:
                if first_cell_text:
                    if is_number(pos_cell.String):
                        align_final = align_final + \
                            "*{" + str(nrcols) + "}{" + dcol_def + "}"
                        break
                    else:
                        alignment_c = pos_cell.HoriJustify
                        if "LEFT" in str(alignment_c):
                            align_final = align_final + "l"
                        elif "CENTER" in str(alignment_c):
                            align_final = align_final + "c"
                        elif "RIGHT" in str(alignment_c):
                            align_final = align_final + "r"
                        elif "STANDARD" in str(alignment_c):
                            align_final = align_final + "c"

    # complete loop for strings
    for i in r_range:
        new_data = []
        for j in c_range:
            o_cell = sheet.getCellByPosition(j, i)
            see_cell = o_cell.String
            see_cell = see_cell.replace("%", "\%")

            # add $...$ to all numeric cells for nice numbers (only if no dc)
            if not dcolumn:
                if is_number(see_cell):
                    see_cell = "$" + see_cell + "$"
                else:
                    pass
            else:
                pass

            # search for italic
            test_italic = o_cell.CharPosture
            searchIT = "com.sun.star.awt.FontSlant ('ITALIC')"

            # search for bold
            test_bold = o_cell.CharWeight
            search_bold = "150"

            # actual loop to replace with latex code
            # Method: nested search for bf and it
            # Loop might be slightly too complicated and ineffective due to
            # iteration for \multicolumn, but it works... (and is still fast
            # enough for any publication-sized table!)
            if search_bold in str(test_bold):
                if searchIT in str(test_italic):
                    it_bf_cell = " \\textit{\\textbf{" + see_cell + "}} "
                    if multicol:
                        if i == frow:
                            it_bf_cell = str_mult_col + it_bf_cell + " } "
                            new_data.append(it_bf_cell)
                        else:
                            new_data.append(it_bf_cell)
                    else:
                        new_data.append(it_bf_cell)
                else:
                    bold_cell = " \\textbf{" + see_cell + "} "
                    if multicol:
                        if i == frow:
                            bold_cell = str_mult_col + bold_cell + "} "
                            new_data.append(bold_cell)
                        else:
                            new_data.append(bold_cell)
                    else:
                        new_data.append(bold_cell)

            elif searchIT in str(test_italic):
                italic_cell = " \\textit{" + see_cell + "} "
                if multicol:
                    if i == frow:
                        italic_cell = str_mult_col + italic_cell + "} "
                        new_data.append(italic_cell)
                    else:
                        new_data.append(italic_cell)
                else:
                    new_data.append(italic_cell)
            else:
                no_cell = " " + see_cell + " "
                if multicol:
                    if i == frow:
                        no_cell = str_mult_col + no_cell + "} "
                        new_data.append(no_cell)
                    else:
                        new_data.append(no_cell)
                else:
                    new_data.append(no_cell)

        data_list.append(new_data)

#################################################################
##############LateX stuff########################################

    lines = []
    for row in data_list:
        lines.append(row)

    # Edit main part of table, add "end-of-line"-characters
    lines_final = []
    for line in lines:
        lastline = line[-1]
        del line[-1]
        newlast = "{0}\\\\".format(lastline)
        line.append(newlast)
        lines_final.append(line)

    # Insert Stuff for booktabs
    lines_final.insert(0, ["\\toprule"])
    lines_final.insert(2, ["\\midrule"])
    lines_final.append(["\\bottomrule"])

    # Define table header
    if not fixed_width:
        header = str_begin_table + '[' + placement + ']\n' + str_centering + \
            '\n\\caption{' + caption + '}\n' + \
            str_begin_tabular + '{' + align_final + '}\n'
    else:
        header = str_begin_table + '[' + placement + ']\n' + str_centering + \
            '\n\\caption{' + caption + '}\n' + str_begin_tabular2 + '{' + \
            str_fixed_width + '\\textwidth}{@{\\extracolsep{\\fill}}' + \
            align_final + '}\n'

    # Write table header
    with open(output, "w") as output_file:
        output_file.write(header)

    # Write main part (body) of table
    with open(output, "a") as output_file:
        lineswrite = csv.writer(output_file, delimiter='&', lineterminator='\n')
        lineswrite.writerows(lines_final)

    # Define footer
    if not fixed_width:
        footer = str_end_tabular + '\n\\label{tab:' + label + '}\n' + str_end_table
    else:
        footer = str_end_tabular2 + '\n\\label{tab:' + label + '}\n' + str_end_table
    # Write footer to file
    with open(output, "a") as output_file:
        output_file.write(footer)

    with open(output, "r") as output_file:
        for line in output_file:
            final_code.append(line)

    for line in final_code:
        final_code_t = final_code_t + line

    os.remove(output)
    
    try:
        # Container / box for all elements
        dialogModel = smgr.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel", ctx)

        dialogModel.PositionX = 50
        dialogModel.PositionY = 50
        dialogModel.Width = 250
        dialogModel.Height = 300
        dialogModel.Title = "LaTeX Table Code ---- calc2booktab"

        # Text-Box to put final code in it...
        textModel = dialogModel.createInstance(
            "com.sun.star.awt.UnoControlEditModel")
        textModel.PositionX = 10
        textModel.PositionY = 10
        textModel.Width = 230
        textModel.Height = 270
        textModel.Name = "lateXCode"
        textModel.Text = str(final_code_t)
        textModel.MultiLine = True
        textModel.ReadOnly = True
        textModel.AutoHScroll = True
        textModel.AutoVScroll = True
        # insert text-box in main container
        dialogModel.insertByName("lateXCode", textModel)

        # select all text in field
        o_selection = Selection(0, len(textModel.Text))

        # initialize container and some other stuff
        controlContainer = smgr.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", ctx)
        controlContainer.setModel(dialogModel)
        toolkit = smgr.createInstanceWithContext(
            "com.sun.star.awt.ExtToolkit", ctx)
        controlContainer.setVisible(True)
        controlContainer.createPeer(toolkit, None)
        oChild = controlContainer.getPeer(
        ).getAccessibleContext().getAccessibleChild(0)
        oChild.setSelection(o_selection)
        controlContainer.execute()
        controlContainer.dispose()
    except Exception(e):
        print(str(e))


# Only export main script to OO-Userinterface. Seems to be buggy at least
# in Windows and openSuse or in older versions of libreoffice... Not used
# until this works.
g_exportedScripts = calc2booktab_basic, \
    calc2booktab_dcolumn, \
    calc2booktab_fixed, \
    calc2booktab_fixed_dcolumn
