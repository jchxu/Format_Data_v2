# coding=utf-8
import xlrd,re
from datetime import datetime

### 获取港口文件名中的简化港口名称,返回字典{港口index:简化名称} ###
def GetShortName(PortFilesList,StdPort):
    Shortnames = {}
    for i in range(0,len(PortFilesList)):
        PortFilename = PortFilesList[i]
        for item in StdPort:
            if item in PortFilename:
                Shortnames[i] = item
                break
    return Shortnames

### 获取各个港口文件名中的日期信息，返回字典{港口index：8位数日期}###
def GetFileDate(PortFilesList,StdPort):
    FileDateDict = {}
    ThisYear = str(datetime.now().year)
    for i in range(0,len(PortFilesList)):
        PortFilename = PortFilesList[i]
        Flag = PortFilename.count('.') - 1
        if Flag == 0:
            FileDate = re.findall(r"\d+", PortFilename)[0]
            if (len(FileDate) == 4): FileDate = ThisYear + FileDate
            elif (len(FileDate) == 6): FileDate = ThisYear[0:2] + FileDate
            FileDateDict[i] = FileDate
        elif Flag == 1:
            TempDate = re.findall(r"\d+", PortFilename)
            FileDate = ThisYear + "%02d"%int(TempDate[0]) + "%02d"%int(TempDate[1])
            FileDateDict[i] = FileDate
        elif Flag == 2:
            TempDate = re.findall(r"\d+", PortFilename)
            if (len(TempDate[0]) == 2): TempDate[0] = ThisYear[0:2] + TempDate[0]
            FileDate = TempDate[0] + "%02d" % int(TempDate[1]) + "%02d" % int(TempDate[2])
            FileDateDict[i] = FileDate
    return (FileDateDict)

### 获取文件名日期 ###
def GetFilename(filename):
    prefix = "铁矿港存结构分析-"
    namelist = ["岚桥", "岚山", "连云港", "京唐港", "实业", "青岛", "日照"]
    flag = filename.count('.') - 1
    std_date = ''
    if flag == 0:
        date = re.findall(r"\d+", filename)[0]
        std_date = date
    elif flag == 1:
        date = re.findall(r"\d+\.?\d*", filename)[0]
        month = int(date.split('.')[0])
        day = int(date.split('.')[1])
        std_date = "%02d%02d" % (month, day)
    elif flag == 2:
        date = re.findall(r"\d+\.?\d+\.?\d*", filename)[0]
        month = int(date.split('.')[-2])
        day = int(date.split('.')[-1])
        std_date = "%02d%02d" % (month, day)
    else:
        print(u"文件名中日期格式不适合，请将日期统一为**.**格式.")
    for item in namelist:
        if item in filename:
            resultname = prefix+item+"-"+std_date
    return resultname

### 读取原始数据，返回列表 ###
def ReadSource(SourceFileName):
    Owner = []
    Goods = []
    Amount = []
    Port = []
    ArrivalDate = []
    SourceFile = xlrd.open_workbook(SourceFileName, 'r')
    Sheet = SourceFile.sheet_by_index(0)
    for i in range(1, Sheet.nrows):
        Line = Sheet.row_values(i)
        Owner.append(Line[0].replace(' ',''))
        Goods.append(Line[1].replace(' ',''))
        Amount.append(Line[2])
        Port.append(Line[3].replace(' ',''))
        if len(Line) >= 5:  #有日期，记录日期；无日期，记录为“-”
            ArrivalDate.append(Line[4])
        else:
            ArrivalDate.append('-')
    print(u'已读取"\033[1;34;0m%s\033[0m"中的\033[1;34;0m%d\033[0m条数据.' % (SourceFileName, Sheet.nrows-1))
    SourceFile.release_resources()
    return (Owner, Goods, Amount, Port, ArrivalDate)

### 读取分类名录中的各个子表，返回为列表，主流粉/块等返回{分类名称：品种}字典
def ReadList(ListFileName):
    GoodsClassName = {}
    GoodsClassList = {}
    ListFile = xlrd.open_workbook(ListFileName, 'r')
    ClassList = ListFile.sheets()[0].col_values(0)  #分类种类列表
    Kinds = ListFile.sheets()[1].col_values(0)  #品种列表
    SteelCompany = ListFile.sheets()[2].col_values(0)   #钢厂列表
    Trader = ListFile.sheets()[3].col_values(0)     #贸易商列表
    for i in range(3, len(ClassList)):
        GoodsClassName[i-3] = ClassList[i].replace(' ','')  #分类种类中的第4项开始为各个小的品种分类
        GoodsClassList[i-3] = ListFile.sheets()[i+1].col_values(0)
    print(u'已读取"\033[1;34;0m%s\033[0m"中的\033[1;34;0m%d\033[0m个清单.' % (ListFileName, len(ClassList)))
    ListFile.release_resources()
    return (Kinds, SteelCompany, Trader, GoodsClassName, GoodsClassList)

### 读取标准名称中的货主和品种标准名称,返回为{一般名称:标准名称}字典 ###
def ReadStd(StdFileName):
    StdOwner = {}
    StdGoods = {}
    StdFile = xlrd.open_workbook(StdFileName, 'r')
    OwnerSheet = StdFile.sheet_by_index(0)
    GoodsSheet = StdFile.sheet_by_index(1)
    for i in range(1, OwnerSheet.nrows):
        RowValue = OwnerSheet.row_values(i)
        StdOwner[RowValue[0].replace(' ','')] = RowValue[1].replace(' ','')
    print(u'已读取"\033[1;34;0m%s\033[0m"中的\033[1;34;0m%d\033[0m个货主标准名称.' % (StdFileName, OwnerSheet.nrows-1))
    for i in range(1, GoodsSheet.nrows):
        RowValue = GoodsSheet.row_values(i)
        StdGoods[RowValue[0].replace(' ','')] = RowValue[1].replace(' ','')
    print(u'已读取"\033[1;34;0m%s\033[0m"中的\033[1;34;0m%d\033[0m个品种标准名称.' % (StdFileName, GoodsSheet.nrows-1))
    StdFile.release_resources()
    return (StdOwner, StdGoods)



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

### 去除列表中各元素或字符串中的空格 ###
def StripSpace(List):
    if (type(List) == list) :
        NewList = []
        for i in range(len(List)):
            if type(List[i]) == str:
                NewList.append(List[i].replace(' ',''))
            else:
                NewList.append(List[i])
        return (NewList)
    elif (type(List) == str) :
        NewStr = ''
        NewStr = List.replace(' ', '')
        return NewStr

### 查找标题行、货主列、品种列、数量列、到港日期列的index ###
def FindIndex(Sheet,OwnWords,GoodWords,AmountWords,DateWords):
    TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = (0,0,0,0,-1)
    for i in range(Sheet.nrows):
        Line = Sheet.row_values(i)
        if ('货主' in Line) or ('收货人' in Line) or ('货名' in Line) :
            TitleRowIndex = i
            LineData = StripSpace(Line)
            for j in range(len(LineData)):
                if LineData[j] in OwnWords: OwnColIndex = j
                elif LineData[j] in GoodWords: GoodColIndex = j
                elif (LineData[j] in AmountWords): AmountColIndex = j
                elif LineData[j] in DateWords: DateColIndex = j
            break
    return (TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex)

def FindStopRow(Sheet, StopWords):
    StopRowIndex = Sheet.nrows
    for i in range(Sheet.nrows):
        Line = Sheet.row_values(i)
        LineData = StripSpace(Line)
        for j in range(len(LineData)):
            LineItem = str(LineData[j])
            for item in StopWords:
                if item in LineItem:
                    StopRowIndex = i + 1
                    return StopRowIndex

### 检查货主或品种是否在合并单元格中 ###
def CheckNameMerge(UnMergeIndexDict, j, OwnColIndex, GoodColIndex):
    OwnRawIndex = j
    GoodRawIndex = j
    OwnIndex = OwnColIndex
    GoodIndex = GoodColIndex
    if (j,OwnColIndex) in UnMergeIndexDict.keys():
        (OwnRawIndex, OwnIndex) = UnMergeIndexDict[(j,OwnColIndex)]
    if (j, GoodColIndex) in UnMergeIndexDict.keys():
        (GoodRawIndex, GoodIndex) = UnMergeIndexDict[(j,GoodColIndex)]
    return (OwnRawIndex, OwnIndex, GoodRawIndex, GoodIndex)

### 检查名称是否为空，或为合计 ###
def CheckFlag(OwnName, GoodName, Amount):
    NameKeyWords = ['总计','合计','统计','商家报告']
    if (StripSpace(GoodName) == ''):
        return False
    for item in NameKeyWords:
        if (item in StripSpace(OwnName)) or (item in StripSpace(GoodName)):
            return False
    if (type(Amount) not in (int,float)):
        return False
    elif (Amount <= 0.0):
        return False
    else:
        return True

### 检查数量是否在合并单元格中 ###
def CheckAmountMerge(UnMergeIndexDict, j, AmountColIndex):
    AmountRawIndex = j
    AmountIndex = AmountColIndex
    if (j,AmountColIndex) in UnMergeIndexDict.keys():
        (AmountRawIndex, AmountIndex) = UnMergeIndexDict[(j,AmountColIndex)]
    return (AmountRawIndex, AmountIndex)

### 检查合并单元格中纵向合并数量 ###
def CountMergeCell(UnMergeIndexDict):
    MergeCount = {}
    Count = {}
    UniqueValue = list(set(list(UnMergeIndexDict.values())))
    for item in UniqueValue:
        Count[item] = 0
    for item in UnMergeIndexDict.keys():
        Count[UnMergeIndexDict[item]] += 1
    for item in list(UnMergeIndexDict.keys()):
        MergeCount[item] = Count[UnMergeIndexDict[item]]
    return (MergeCount)

### 返回合并单元格纵向数量，若非则返回1 ###
def CheckAmountRatio(MergeCount, AmountRawIndex, AmountIndex):
    AmountRatio = 1
    if (AmountRawIndex, AmountIndex) in MergeCount.keys():
        AmountRatio = MergeCount[(AmountRawIndex, AmountIndex)]
    return (AmountRatio)

### 格式化日期为yyyymmdd ###
def FormatDate(DateStr, DateCtype):
    FormatDate = ''
    if DateCtype == 1:
        if '.' in DateStr:
            Flag = (DateStr).count('.')
            if Flag == 2:  #yy.mm.dd
                year = DateStr.split('.')[0]
                month = DateStr.split('.')[1]
                day = DateStr.split('.')[2]
                if (len(year) == 2): year = '20' + year
                if (len(month) == 1): month = '0' + month
                if (len(day) == 1): day = '0' + day
                FormatDate = year+month+day
    elif DateCtype == 3:
        TempDate = xlrd.xldate_as_tuple(DateStr,0)
        year = str(TempDate[0])
        month = str(TempDate[1])
        day = str(TempDate[2])
        if (len(year) == 2): year = '20' + year
        if (len(month) == 1): month = '0' + month
        if (len(day) == 1): day = '0' + day
        FormatDate = year + month + day
    return FormatDate

### 读取曹妃甸(实业)港口数据 ###
def ReadPort01(PortFilename,PortDate):
    OwnWords = ['货主']
    GoodWords = ['货种']
    AmountWords = ['场存数量']
    DateWords = ['入库时间']
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)
    Sheets = PortFile.sheets()
    FileData = []
    for i in range(len(Sheets)):
        Sheet = PortFile.sheet_by_index(i)
        UnMergeIndexDict = UnMergeCell(Sheet.merged_cells)
        MergeCount = CountMergeCell(UnMergeIndexDict)
        TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = FindIndex(Sheet,OwnWords,GoodWords,AmountWords,DateWords)
        for j in range(TitleRowIndex + 1, Sheet.nrows):
            OwnRawIndex, OwnIndex, GoodRawIndex, GoodIndex = CheckNameMerge(UnMergeIndexDict, j, OwnColIndex, GoodColIndex)
            OwnName = StripSpace(Sheet.cell_value(OwnRawIndex,OwnIndex))
            GoodName = StripSpace(Sheet.cell_value(GoodRawIndex,GoodIndex))
            AmountRawIndex, AmountIndex = CheckAmountMerge(UnMergeIndexDict, j, AmountColIndex)
            AmountRatio = CheckAmountRatio(MergeCount, AmountRawIndex, AmountIndex)
            Amount = StripSpace(Sheet.row_values(AmountRawIndex))[AmountIndex] * AmountRatio
            DateCtype = Sheet.cell(AmountRawIndex,DateColIndex).ctype
            RecDate = FormatDate(Sheet.cell_value(AmountRawIndex,DateColIndex),DateCtype)
            Flag = CheckFlag(OwnName, GoodName, Amount)
            if Flag:
                FileData.append([AmountRawIndex+1, OwnName, GoodName, Amount, RecDate])
    PortFile.release_resources()
    return (FileData)

### 读取曹妃甸(弘毅)港口数据 ###
def ReadPort02(PortFilename,PortDate):
    OwnWords = ['货主']
    GoodWords = ['品名']
    AmountWords = ['结存']
    DateWords = ['到港日期']
    StopWords = ['集港车数']
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)
    Sheets = PortFile.sheets()
    FileData = []
    for i in range(len(Sheets)):
        Sheet = PortFile.sheet_by_index(i)
        if Sheet.name == '散货库存':
            UnMergeIndexDict = UnMergeCell(Sheet.merged_cells)
            MergeCount = CountMergeCell(UnMergeIndexDict)
            TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = FindIndex(Sheet,OwnWords,GoodWords,AmountWords,DateWords)
            StopRowIndex = FindStopRow(Sheet, StopWords)
            for j in range(TitleRowIndex + 1, StopRowIndex):
                OwnRawIndex, OwnIndex, GoodRawIndex, GoodIndex = CheckNameMerge(UnMergeIndexDict, j, OwnColIndex, GoodColIndex)
                OwnName = StripSpace(Sheet.cell_value(OwnRawIndex,OwnIndex))
                GoodName = StripSpace(Sheet.cell_value(GoodRawIndex,GoodIndex))
                AmountRawIndex, AmountIndex = CheckAmountMerge(UnMergeIndexDict, j, AmountColIndex)
                AmountRatio = CheckAmountRatio(MergeCount, AmountRawIndex, AmountIndex)
                Amount = StripSpace(Sheet.row_values(AmountRawIndex))[AmountIndex] * AmountRatio
                DateCtype = Sheet.cell(AmountRawIndex,DateColIndex).ctype
                RecDate = FormatDate(Sheet.cell_value(AmountRawIndex,DateColIndex),DateCtype)
                Flag = CheckFlag(OwnName, GoodName, Amount)
                if Flag:
                    FileData.append([AmountRawIndex+1, OwnName, GoodName, Amount, RecDate])
    PortFile.release_resources()
    return (FileData)

### 读取曹妃甸(矿三)港口数据 ###
def ReadPort03(PortFilename,PortDate):
    OwnWords = ['货主']
    GoodWords = ['货种']
    AmountWords = ['结存（吨）']
    DateWords = ['入库日期']
    StopWords = ['废矿']
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)
    Sheets = PortFile.sheets()
    FileData = []
    #for i in range(len(Sheets)):
    Sheet = PortFile.sheet_by_index(0)
    UnMergeIndexDict = UnMergeCell(Sheet.merged_cells)
    MergeCount = CountMergeCell(UnMergeIndexDict)
    TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = FindIndex(Sheet, OwnWords,GoodWords, AmountWords,DateWords)
    StopRowIndex = FindStopRow(Sheet, StopWords)
    for j in range(TitleRowIndex + 1, StopRowIndex):
        OwnRawIndex, OwnIndex, GoodRawIndex, GoodIndex = CheckNameMerge(UnMergeIndexDict, j, OwnColIndex,GoodColIndex)
        OwnName = StripSpace(Sheet.cell_value(OwnRawIndex, OwnIndex))
        GoodName = StripSpace(Sheet.cell_value(GoodRawIndex, GoodIndex))
        AmountRawIndex, AmountIndex = CheckAmountMerge(UnMergeIndexDict, j, AmountColIndex)
        AmountRatio = CheckAmountRatio(MergeCount, AmountRawIndex, AmountIndex)
        Amount = StripSpace(Sheet.row_values(AmountRawIndex))[AmountIndex] * AmountRatio
        DateCtype = Sheet.cell(AmountRawIndex, DateColIndex).ctype
        RecDate = FormatDate(Sheet.cell_value(AmountRawIndex, DateColIndex), DateCtype)
        Flag = CheckFlag(OwnName, GoodName, Amount)
        if Flag:
            FileData.append([AmountRawIndex+1, OwnName, GoodName, Amount, RecDate])
    PortFile.release_resources()
    return (FileData)

### 读取京唐港口数据 ###
def ReadPort1(PortFilename,PortDate):
    OwnWords = ['货主']
    GoodWords = ['货名']
    AmountWords = ['结存量']
    DateWords = ['入库时间']
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)
    Sheets = PortFile.sheets()
    FileData = []
    for i in range(len(Sheets)):
        Sheet = PortFile.sheet_by_index(i)
        UnMergeIndexDict = UnMergeCell(Sheet.merged_cells)
        MergeCount = CountMergeCell(UnMergeIndexDict)
        TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = FindIndex(Sheet, OwnWords, GoodWords,AmountWords, DateWords)
        for j in range(TitleRowIndex + 1, Sheet.nrows):
            OwnRawIndex, OwnIndex, GoodRawIndex, GoodIndex = CheckNameMerge(UnMergeIndexDict, j, OwnColIndex,GoodColIndex)
            OwnName = StripSpace(Sheet.cell_value(OwnRawIndex, OwnIndex))
            GoodName = StripSpace(Sheet.cell_value(GoodRawIndex, GoodIndex))
            AmountRawIndex, AmountIndex = CheckAmountMerge(UnMergeIndexDict, j, AmountColIndex)
            AmountRatio = CheckAmountRatio(MergeCount, AmountRawIndex, AmountIndex)
            Amount = StripSpace(Sheet.row_values(AmountRawIndex))[AmountIndex] * AmountRatio
            DateCtype = Sheet.cell(AmountRawIndex, DateColIndex).ctype
            RecDate = FormatDate(Sheet.cell_value(AmountRawIndex, DateColIndex), DateCtype)
            Flag = CheckFlag(OwnName, GoodName, Amount)
            if Flag:
                FileData.append([AmountRawIndex+1, OwnName, GoodName, Amount, RecDate])
    PortFile.release_resources()
    return (FileData)

### 读取岚桥港口数据 ###
def ReadPort2(PortFilename,PortDate):
    OwnWords = ['货主']
    GoodWords = ['货性']
    AmountWords = ['剩余货权量（吨）']
    DateWords = ['到港日期']
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)
    Sheets = PortFile.sheets()
    FileData = []
    for i in range(len(Sheets)):
        Sheet = PortFile.sheet_by_index(i)
        UnMergeIndexDict = UnMergeCell(Sheet.merged_cells)
        MergeCount = CountMergeCell(UnMergeIndexDict)
        TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = FindIndex(Sheet,OwnWords,GoodWords,AmountWords,DateWords)
        for j in range(TitleRowIndex + 1, Sheet.nrows):
            OwnRawIndex, OwnIndex, GoodRawIndex, GoodIndex = CheckNameMerge(UnMergeIndexDict, j, OwnColIndex, GoodColIndex)
            OwnName = StripSpace(Sheet.cell_value(OwnRawIndex, OwnIndex))
            GoodName = StripSpace(Sheet.cell_value(GoodRawIndex, GoodIndex))
            AmountRawIndex, AmountIndex = CheckAmountMerge(UnMergeIndexDict, j, AmountColIndex)
            AmountRatio = CheckAmountRatio(MergeCount, AmountRawIndex, AmountIndex)
            Amount = StripSpace(Sheet.row_values(AmountRawIndex))[AmountIndex] * AmountRatio
            Flag = CheckFlag(OwnName, GoodName, Amount)
            if Flag:
                FileData.append([AmountRawIndex+1, OwnName, GoodName, Amount, PortDate])
    PortFile.release_resources()
    return (FileData)

### 读取岚山港口数据 ###
def ReadPort3(PortFilename,PortDate):
    OwnWords = ['货主']
    GoodWords = ['货性']
    AmountWords = ['数量（吨）']
    DateWords = ['到港日期']
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)
    Sheets = PortFile.sheets()
    FileData = []
    for i in range(len(Sheets)):
        Sheet = PortFile.sheet_by_index(i)
        UnMergeIndexDict = UnMergeCell(Sheet.merged_cells)
        MergeCount = CountMergeCell(UnMergeIndexDict)
        TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = FindIndex(Sheet,OwnWords,GoodWords,AmountWords,DateWords)
        for j in range(TitleRowIndex + 1, Sheet.nrows):
            OwnRawIndex, OwnIndex, GoodRawIndex, GoodIndex = CheckNameMerge(UnMergeIndexDict, j, OwnColIndex, GoodColIndex)
            OwnName = StripSpace(Sheet.cell_value(OwnRawIndex, OwnIndex))
            GoodName = StripSpace(Sheet.cell_value(GoodRawIndex, GoodIndex))
            AmountRawIndex, AmountIndex = CheckAmountMerge(UnMergeIndexDict, j, AmountColIndex)
            AmountRatio = CheckAmountRatio(MergeCount, AmountRawIndex, AmountIndex)
            Amount = StripSpace(Sheet.row_values(AmountRawIndex))[AmountIndex] * AmountRatio
            Flag = CheckFlag(OwnName, GoodName, Amount)
            if Flag:
                FileData.append([AmountRawIndex+1, OwnName, GoodName, Amount, PortDate])
    PortFile.release_resources()
    return (FileData)

### 读取连云港港口数据 ###
def ReadPort4(PortFilename,PortDate):
    OwnWords = ['货主','钢厂及']
    GoodWords = ['货名','品种']
    AmountWords = ['数量/吨','港存数','数量','重量']
    DateWords = ['到港船期','靠港时间','卸船日期','进场日期']
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)
    Sheets = PortFile.sheets()
    FileData = []
    for i in range(len(Sheets)):
        Sheet = PortFile.sheet_by_index(i)
        UnMergeIndexDict = UnMergeCell(Sheet.merged_cells)
        MergeCount = CountMergeCell(UnMergeIndexDict)
        TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = FindIndex(Sheet, OwnWords, GoodWords,AmountWords, DateWords)
        for j in range(TitleRowIndex + 1, Sheet.nrows):
            OwnRawIndex, OwnIndex, GoodRawIndex, GoodIndex = CheckNameMerge(UnMergeIndexDict, j, OwnColIndex,GoodColIndex)
            OwnName = StripSpace(Sheet.cell_value(OwnRawIndex, OwnIndex))
            GoodName = StripSpace(Sheet.cell_value(GoodRawIndex, GoodIndex))
            AmountRawIndex, AmountIndex = CheckAmountMerge(UnMergeIndexDict, j, AmountColIndex)
            AmountRatio = CheckAmountRatio(MergeCount, AmountRawIndex, AmountIndex)
            Amount = StripSpace(Sheet.row_values(AmountRawIndex))[AmountIndex] * AmountRatio
            DateCtype = Sheet.cell(AmountRawIndex, DateColIndex).ctype
            RecDate = FormatDate(Sheet.cell_value(AmountRawIndex, DateColIndex), DateCtype)
            Flag = CheckFlag(OwnName, GoodName, Amount)
            if Flag:
                FileData.append([AmountRawIndex+1, OwnName, GoodName, Amount, RecDate])
    PortFile.release_resources()
    return (FileData)

### 读取青岛港口数据 ###
def ReadPort5(PortFilename,PortDate):
    OwnWords = ['收货人']
    GoodWords = ['货名']
    AmountWords = ['全船结存']
    DateWords = ['卸船日期']
    StopWords = ['分类统计']
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)
    Sheets = PortFile.sheets()
    FileData = []
    # for i in range(len(Sheets)):
    Sheet = PortFile.sheet_by_index(0)
    UnMergeIndexDict = UnMergeCell(Sheet.merged_cells)
    MergeCount = CountMergeCell(UnMergeIndexDict)
    TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = FindIndex(Sheet, OwnWords, GoodWords,AmountWords, DateWords)
    StopRowIndex = FindStopRow(Sheet, StopWords)
    for j in range(TitleRowIndex + 1, StopRowIndex):
        OwnRawIndex, OwnIndex, GoodRawIndex, GoodIndex = CheckNameMerge(UnMergeIndexDict, j, OwnColIndex, GoodColIndex)
        OwnName = StripSpace(Sheet.cell_value(OwnRawIndex, OwnIndex))
        GoodName = StripSpace(Sheet.cell_value(GoodRawIndex, GoodIndex))
        AmountRawIndex, AmountIndex = CheckAmountMerge(UnMergeIndexDict, j, AmountColIndex)
        AmountRatio = CheckAmountRatio(MergeCount, AmountRawIndex, AmountIndex)
        Amount = StripSpace(Sheet.row_values(AmountRawIndex))[AmountIndex] * AmountRatio
        DateCtype = Sheet.cell(AmountRawIndex, DateColIndex).ctype
        RecDate = FormatDate(Sheet.cell_value(AmountRawIndex, DateColIndex), DateCtype)
        Flag = CheckFlag(OwnName, GoodName, Amount)
        if Flag:
            FileData.append([AmountRawIndex + 1, OwnName, GoodName, Amount, RecDate])
    PortFile.release_resources()
    return (FileData)

### 读取日照港口数据 ###
def ReadPort6(PortFilename,PortDate):
    OwnWords = ['收货人（货主）']
    GoodWords = ['货名']
    AmountWords = ['当日库存']
    DateWords = ['卸船日期']
    PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)
    Sheets = PortFile.sheets()
    FileData = []
    # for i in range(len(Sheets)):
    Sheet = PortFile.sheet_by_index(0)
    UnMergeIndexDict = UnMergeCell(Sheet.merged_cells)
    MergeCount = CountMergeCell(UnMergeIndexDict)
    TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = FindIndex(Sheet, OwnWords, GoodWords,AmountWords, DateWords)
    for j in range(TitleRowIndex + 1, Sheet.nrows):
        OwnRawIndex, OwnIndex, GoodRawIndex, GoodIndex = CheckNameMerge(UnMergeIndexDict, j, OwnColIndex, GoodColIndex)
        OwnName = StripSpace(Sheet.cell_value(OwnRawIndex, OwnIndex))
        GoodName = StripSpace(Sheet.cell_value(GoodRawIndex, GoodIndex))
        AmountRawIndex, AmountIndex = CheckAmountMerge(UnMergeIndexDict, j, AmountColIndex)
        AmountRatio = CheckAmountRatio(MergeCount, AmountRawIndex, AmountIndex)
        Amount = StripSpace(Sheet.row_values(AmountRawIndex))[AmountIndex] * AmountRatio
        DateCtype = Sheet.cell(AmountRawIndex, DateColIndex).ctype
        RecDate = FormatDate(Sheet.cell_value(AmountRawIndex, DateColIndex), DateCtype)
        Flag = CheckFlag(OwnName, GoodName, Amount)
        if Flag:
            FileData.append([AmountRawIndex + 1, OwnName, GoodName, Amount, RecDate])
    PortFile.release_resources()
    return (FileData)

### 根据港口名称读取数据 ###
#'曹妃甸','京唐','岚桥','岚山','连云港','青岛','日照'
def ReadData(PortFilename,PortShortname,PortDate):
    PortData = []
    if PortShortname == '曹妃甸实业': PortData = ReadPort01(PortFilename,PortDate)
    elif PortShortname == '曹妃甸弘毅': PortData = ReadPort02(PortFilename,PortDate)
    elif PortShortname == '曹妃甸矿三': PortData = ReadPort03(PortFilename,PortDate)
    elif PortShortname == '京唐': PortData = ReadPort1(PortFilename,PortDate)
    elif PortShortname == '岚桥': PortData = ReadPort2(PortFilename,PortDate)
    elif PortShortname == '岚山': PortData = ReadPort3(PortFilename,PortDate)
    elif PortShortname == '连云港': PortData = ReadPort4(PortFilename,PortDate)
    elif PortShortname == '青岛': PortData = ReadPort5(PortFilename,PortDate)
    elif PortShortname == '日照': PortData = ReadPort6(PortFilename,PortDate)
    else: print(PortFilename,'港口名称或港口名称不在标准名称中')
    return PortData

### 检查吨/万吨？ ###
def CheckTons(PortData):
    ref = 50
    templist = []
    for item in PortData:
        if item[3] <= ref:
            templist.append(item[3])
    AllNum = len(PortData)
    SmallNum = len(templist)
    #print(SmallNum,AllNum)
    if SmallNum > 0.99*AllNum:
        for item in PortData:
            item[3] = item[3]*10000
    return PortData

### 读取各港口数据 ###
def CollectPortData(PortFilesList,StdPort):
    AllData = []    #最终存储的数据，[货主，品种，库存，港口，到港日期]
    PortDates = GetFileDate(PortFilesList, StdPort)
    PortShortnames = GetShortName(PortFilesList,StdPort)
    for i in range(0,len(PortFilesList)):
        PortFilename = PortFilesList[i]
        PortData = ReadData(PortFilename,PortShortnames[i],PortDates[i])
        PortData = CheckTons(PortData)
        for item in PortData: item.append(PortShortnames[i])
        print('已读取"%s"文件中的%d条数据.' % (PortFilename,len(PortData)))
        #for item in PortData: print(item)
        AllData += PortData
    return (AllData)