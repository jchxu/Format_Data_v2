# coding=utf-8
import xlrd

### 返回字典{合并的单元格坐标:合并单元格左上角坐标(即合并前数值所在单元格的坐标)}
def UnMergeCell(MergedCellsList):
    UnMergeIndexDict = {}
    for item in MergedCellsList:
        rowlow = item[0]
        rowhigh = item[1]
        collow = item[2]
        colhigh = item[3]
        for i in range(rowlow,rowhigh):
            for j in range(collow,colhigh):
                UnMergeIndexDict[(i,j)] = (rowlow,collow)
    return (UnMergeIndexDict)

### 去除列表中各元素中的空格 ###
def StripSpace(List):
    NewList = []
    for i in range(len(List)):
        NewList.append(List[i].replace(' ',''))
    return (NewList)

### 查找标题行、货主列、品种列、数量列、到港日期列的index ###
def FindIndex(Sheet):
    OwnWords = ['货主', '收货人']
    GoodWords = ['货性']
    AmountWords = ['数量','数量（吨）']
    DateWords = []
    TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = (0,0,0,0,-1)
    for i in range(Sheet.nrows):
        Line = Sheet.row_values(i)
        if ('货主' in Line) or ('收货人' in Line) :
            TitleRowIndex = i
            print(Sheet.row_values(i))
            LineData = StripSpace(Line)
            for j in range(len(LineData)):
                if LineData[j] in OwnWords: OwnColIndex = j
                elif LineData[j] in GoodWords: GoodColIndex = j
                elif LineData[j] in AmountWords: AmountColIndex = j
                elif LineData[j] in DateWords: DateColIndex = j
            break
    print(TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex)
    return (TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex)

### 检查DateIndex是否有效 ###
def CheckDateIndex(OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex):
    if (OwnColIndex != DateColIndex) and (GoodColIndex != DateColIndex) and (AmountColIndex != DateColIndex) and (DateColIndex != -1): return True
    else: return False









PortFilename = '连云港港存贸易矿8.12.xls'
PortFilename = '岚山现货表 08.14.xls'
PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)
ProDate = 20180814

Sheets = PortFile.sheets()
for i in range(len(Sheets)):
    Sheet = PortFile.sheet_by_index(i)
    UnMergeIndexDict = UnMergeCell(Sheet.merged_cells)
    TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = FindIndex(Sheet)
    CheckDateIndex(OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex)

    for j in range(TitleRowIndex+1,Sheet.nrows):
        LineData = Sheet.row_values(j)
        CheckNameMerge(j,OwnColIndex, GoodColIndex)
        CheckAmountMerge(j,OwnColIndex, GoodColIndex)

        print(LineData[OwnColIndex],LineData[GoodColIndex],LineData[AmountColIndex])



PortFile.release_resources()