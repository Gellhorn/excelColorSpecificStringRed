from xlsxwriter.workbook import Workbook

workbook = Workbook('sequences.xlsx')
worksheet = workbook.add_worksheet()

red = workbook.add_format({'color': 'red'})
green = workbook.add_format({'color': 'green'})

sequences = [
    'BCBBGBTG',
    'CCATTGTC',
    'CCCCGGCC',
    'CCTGCTGC',
    'GCTGCTCT',
    'CGGGGCCA',
    'GGCCACCG',
]



for row_num, sequence in enumerate(sequences):

    format_pairs = []

    # Get each DNA base character from the sequence.
    for base in sequence:

        # Prefix each base with a format.
        if base == 'A':
            format_pairs.extend((red, base))
        else:
            format_pairs.append(base)

    worksheet.write_rich_string(row_num, 0, *format_pairs)

workbook.close()
