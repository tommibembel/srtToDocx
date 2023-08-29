import io
import re
import datetime
import docx
from docx.shared import RGBColor
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.enum.table import WD_TABLE_ALIGNMENT


def convert(inputfile, outputfile):
    if inputfile.endswith('.srt'):
        f = io.open(inputfile, encoding="utf-8-sig", mode="r")
        time_re = re.compile(r'\d{2}:\d{2}:\d{2},\d{3} --> \d{2}:\d{2}:\d{2},\d{3}')
        result = {}
        key = ""
        subtitle = ""
        for line in f.readlines():
            if line.strip().isdigit():
                key = line

            elif re.search(time_re, line):
                times = line.strip().split(' --> ')
                dt_start = datetime.datetime.strptime(times[0], "%H:%M:%S,%f")
                dt_end = datetime.datetime.strptime(times[1], "%H:%M:%S,%f")
                dt_dur = dt_end - dt_start

            elif line in ["\n", "\r\n"] and not subtitle == "" and not key == "":
                result[int(key)] = (
                dt_start.strftime("%H:%M:%S,%f")[:-3], dt_end.strftime("%H:%M:%S,%f")[:-3], str(dt_dur)[:-3], subtitle)
                subtitle = ""
                key = ""

            elif line in ["\n", "\r\n"]:
                print("Error in line")

            else:
                subtitle += line

        if subtitle != "" and key != "":
            result[int(key)] = (
                dt_start.strftime("%H:%M:%S,%f")[:-3], dt_end.strftime("%H:%M:%S,%f")[:-3], str(dt_dur)[:-3], subtitle)

        document = docx.Document()
        table = document.add_table(rows=1, cols=5, style='Table Grid')
        table.allow_autofit = True
        table.alignment = WD_TABLE_ALIGNMENT.RIGHT

        tableheaders = table.rows[0].cells
        tableheaders[0].text = '#'
        tableheaders[1].text = 'Start Time'
        tableheaders[2].text = 'End Time'
        tableheaders[3].text = 'Duration'
        tableheaders[4].text = 'Text'

        # Set fontcolor to black and bg color of to white for table headers
        for i in range(len(tableheaders)):
            tableheaders[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
            bgc = parse_xml(r'<w:shd {} w:fill="000000"/>'.format(nsdecls('w')))
            table.rows[0].cells[i]._tc.get_or_add_tcPr().append(bgc)

        for index, (startTime, endTime, duration, subtitle) in sorted(result.items()):
            row_cells = table.add_row().cells
            row_cells[0].text = str(index)
            row_cells[1].text = str(startTime)
            row_cells[2].text = str(endTime)
            row_cells[3].text = str(duration)
            if subtitle.find("<i>") >= 0:
                p = row_cells[4].paragraphs[0].insert_paragraph_before(text=None, style=None)
                for line in subtitle.removesuffix("\n").splitlines(True):
                    if line.find("<i>") >= 0:
                        index1 = line.index("<i>")
                        index2 = line.index("</i>")
                        before = line[0:index1]
                        after = line[index2 + 4:]
                        italic_text = line[index1 + 3:index2]
                        p.add_run(before)
                        run = p.add_run(italic_text)
                        run.italic = True
                        p.add_run(after)
            else:
                row_cells[4].text = str(subtitle)

        for cell in table.columns[0].cells:
            cell.width = 0.5 * 914400

        for cell in table.columns[4].cells:
            cell.width = 4 * 914400
        try:
            document.save(outputfile)
            return 0
        except Exception as e:
            return e

