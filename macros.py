# Python macros for LibreOffice Calc.

def hello_calc(*args):
    """Prints python info into the first cells of current sheet in Calc."""
    document = XSCRIPTCONTEXT.getDocument()
    sheet = document.CurrentController.ActiveSheet
    import sys
    sheet.getCellByPosition(0, 0).String = "Python version is %s.%s.%s" % sys.version_info[:3]
    sheet.getCellByPosition(0, 1).String = "Executable path is " + sys.executable

def analyze_reel(*args):
    """Read a number from the first cell as slot game reel data. Print analysis table below it."""

    document = XSCRIPTCONTEXT.getDocument()
    sheet = document.CurrentController.ActiveSheet

    reel = sheet.getCellByPosition(0, 0).String
    if reel == "" or len(reel) < 3:
        sheet.getCellByPosition(0, 0).String = "put number here"
        return

    # List windows of 3 reel symbols.
    trios = []
    for i in range(len(reel)):
        end = i + 3
        if end < len(reel):
            trios.append(reel[i:end])
        else:
            end = end - len(reel)
            trios.append(reel[i:] + reel[:end])

    # Count symbol appearances within windows.
    counts = []
    for i in range(10):
        counts.append([])
        for trio in trios:
            counts[i].append([])
            counts[i][-1] = len( trio.split(str(i)) ) - 1

    # Print.
    sheet.getCellRangeByPosition(0, 1, len(trios) + 1, len(counts) + 1).clearContents(1 + 4)
    for i in range(len(trios)):
        sheet.getCellByPosition(1 + i, 1).String = trios[i]
    for i in range(len(counts)):
        sheet.getCellByPosition(0, 2 + i).String = "count " + str(i)
    sheet.getCellRangeByPosition(1, 2, len(trios), len(counts) + 1).setDataArray(counts)
