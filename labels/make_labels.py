#! /usr/bin/python

import xlrd

tex_header = r"""
\documentclass[letterpaper,10pt]{article}

\usepackage[document]{ragged2e}
\usepackage{multicol}
\usepackage{textcomp}
\usepackage[cm]{fullpage}

\begin{document}
\pagenumbering{gobble}
"""

tex_footer = """

\end{document}
"""


workbook = xlrd.open_workbook('../collections.xlsx')
worksheet = workbook.sheet_by_name('Sheet1')
tex = ""
odd = True
for row in range(worksheet.nrows):
    if row != 0 and worksheet.cell_value(row, 6) == 'Hesperolinon':
        collection_id = str(int(worksheet.cell_value(row, 0)))
        date = xlrd.xldate_as_tuple(worksheet.cell_value(row, 2), workbook.datemode)
        day = str(date[2])
        species = worksheet.cell_value(row, 6) + ' ' + worksheet.cell_value(row, 7)
        author = worksheet.cell_value(row, 9)
        latitude = str(worksheet.cell_value(row, 10))
        longitude = str(worksheet.cell_value(row, 11))
        elevation_ft = worksheet.cell_value(row, 12)
        elevation = str(int(float(elevation_ft)*0.3048))
        place = worksheet.cell_value(row, 13)
        place = place.replace("Co., CA", "County, California")
        locality = worksheet.cell_value(row, 14)
        habitat = worksheet.cell_value(row, 15)
        yuri_pop = worksheet.cell_value(row, 16)
        if yuri_pop.find("#") != -1:
            habitat += r""" Y.P. Springer (2007) population \#""" + yuri_pop[yuri_pop.find("#")+1:yuri_pop.find("#")+5] + "."

        if odd:
            tex += r"""
%
% row begin
%
"""
        # start the label
        tex += r"""
% label start
\begin{minipage}[t]{0.40\textwidth}

\begin{center}
University of California Herbarium \\
\begin{large}
Plants of """ + place + r""" \\
\end{large}
\vspace{\baselineskip}
\textbf{Linaceae} \\
\textit{""" + species + r"""} """ + author + r"""\\
\end{center}

\begin{footnotesize}

\begin{multicols}{2}
""" + latitude[:6] + r"""\textdegree N """ + longitude[:7] + r"""\textdegree W
\columnbreak
\begin{flushright}
Elev. """ + elevation + r""" m
\end{flushright}
\end{multicols}

""" + locality + r"""
\vspace{\baselineskip}

""" + habitat + r"""

\begin{multicols}{2}
W.A. Freyman """ + collection_id + r""" \\
with A.C. Schneider \& Y.P. Springer
\columnbreak
\begin{flushright}
""" + str(day) + r""" May 2014
\end{flushright}
\end{multicols}

\end{footnotesize}

\end{minipage}
% label end
"""

        if odd:
            odd = False
            tex += r"""%
\hspace{2cm}
%"""
        else:
            odd = True
            tex += r"""
\vspace{2cm}
%
% row end
%
"""

tex = tex_header + tex + tex_footer

tex_file = open("hesperolinon_labels.tex", "w")
tex_file.write(tex)



