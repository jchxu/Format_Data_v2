# coding=utf-8
import xlsxwriter,datetime

### 计算百分比，返回*100的结果，总数为0则全部返回0 ###
def CalcRatio(Total, Steel, Trade):
    SteelRatio = 0
    TradeRatio = 0
    if Total != 0:
        SteelRatio = 1 * Steel / Total
        TradeRatio = 1 * Trade / Total
    return (SteelRatio, TradeRatio)

### TopShip格式化： 货主(数量)，货主(数量)……###
def TopShipFormat(TopList):
    SepStr = '/'
    List = []
    for i in range(0,len(TopList)):
        List.append('%s(%.0f)' % (TopList[i][0],TopList[i][1]))
    return (SepStr.join(List))

### 统计货权排名（按数量统计排名） ###
def TopShip(GoodShipDict):
    GoodTop13 = []
    GoodTop46 = []
    GoodTopOther = []
    Sortedlist = sorted(GoodShipDict.items(),key = lambda x:x[1],reverse = True)
    if len(Sortedlist) <= 3:
        GoodTop13 = Sortedlist
    elif 3 < len(Sortedlist) <= 6:
        GoodTop13 = Sortedlist[0:3]
        GoodTop46 = Sortedlist[3:6]
    elif len(Sortedlist) > 6:
        GoodTop13 = Sortedlist[0:3]
        GoodTop46 = Sortedlist[3:6]
        GoodTopOther = Sortedlist[6:len(Sortedlist)]
    GoodTop13 = TopShipFormat(GoodTop13)
    GoodTop46 = TopShipFormat(GoodTop46)
    GoodTopOther = TopShipFormat(GoodTopOther)
    return (GoodTop13,GoodTop46,GoodTopOther)

### 计算不同分类的行index，假设只列出前两类的明细 ###
def ClassLineIndex(GoodsClassName, GoodsClassList):
    ClassLineIndexDict = {}
    ClassLineIndexDict[0] = 0
    for i in range(1,len(GoodsClassName)):
        ClassLineIndexDict[i] = len(GoodsClassList[i-1]) + ClassLineIndexDict[i-1] + 1
        #if i >= 3: ClassLineIndexDict[i] = ClassLineIndexDict[i-1] + 1  #假设只列出前两类的明细
    return ClassLineIndexDict


### 打印输出各港口、分类、品种汇总数量、货权集中度 （屏幕输出、txt版）###
#AmountInfo = [TotalAmount, TotalSteel, ClassTotal, ClassSteel, GoodsTotal, GoodsSteel]
#ShipInfo = [GoodShip, GoodSteelShip, GoodOtherShip]
def WriteTXT(ResFileName,AmountInfo,ShipInfo, PortList, GoodsClassName, GoodsClassList):
    for i in range(0,len(PortList)):
        print(PortList[i])
        port = PortList[i]
        totalamount = AmountInfo[0][port]
        totalsteel = AmountInfo[1][port]
        totaltrade = totalamount - totalsteel
        steeltotalratio, tradetotalratio = CalcRatio(totalamount, totalsteel, totaltrade)
        print("%s,%.0f,%.0f,%.1f%%,%.0f,%.1f%%" % ('合计', totalamount, totalsteel, steeltotalratio, totaltrade, tradetotalratio), sep=',')
        for j in range(0,len(GoodsClassName)):
            classname = GoodsClassName[j]
            calsstotal = AmountInfo[2][port][classname]
            classsteel = AmountInfo[3][port][classname]
            classtrade = calsstotal - classsteel
            steelratio, traderatio = CalcRatio(calsstotal, classsteel, classtrade)
            print("%s,%.0f,%.0f,%.1f%%,%.0f,%.1f%%" % (classname,calsstotal,classsteel, steelratio, classtrade, traderatio), sep=',')
            if (j == 0) or (j == 1): #pass
                for k in range(0,len(GoodsClassList[j])):
                    goodname = GoodsClassList[j][k]
                    if goodname in AmountInfo[4][port].keys():
                        goodtotal = AmountInfo[4][port][goodname]
                        goodsteel = AmountInfo[5][port][goodname]
                        goodtrade = goodtotal - goodsteel
                        goodship = ShipInfo[0][port][goodname]
                        goodsteelship = ShipInfo[1][port][goodname]
                        goodothership = ShipInfo[2][port][goodname]
                        GoodTop13, GoodTop46, GoodTopOther = TopShip(goodship)
                        GoodSteelTop13, GoodSteelTop46, GoodSteelTopOther = TopShip(goodsteelship)
                        GoodOtherTop13, GoodOtherTop46, GoodOtherTopOther = TopShip(goodothership)
                    else:
                        goodtotal = 0
                        goodsteel = 0
                        goodtrade = 0
                        GoodTop13, GoodTop46, GoodTopOther = ('','','')
                    steelratio, traderatio = CalcRatio(goodtotal, goodsteel, goodtrade)
                    print("%s,%.0f,%.0f,%.1f%%,%.0f,%.1f%%" % (goodname, goodtotal, goodsteel, steelratio, goodtrade, traderatio), sep=',', end = ',')
                    print("%s,%s,%s,%s,%s,%s,%s,%s,%s" % (GoodTop13, GoodTop46, GoodTopOther,GoodSteelTop13, GoodSteelTop46, GoodSteelTopOther,GoodOtherTop13, GoodOtherTop46, GoodOtherTopOther), sep=',')

### 打印输出各港口、分类、品种汇总数量、货权集中度 （Excel版）###
def WriteExcel(ResFileName,AmountInfo,ShipInfo, PortList, GoodsClassName, GoodsClassList):
    Filename = ResFileName+'.xlsx'
    WorkBook = xlsxwriter.Workbook(Filename)
    Sheet1 = WorkBook.add_worksheet()
    Sheet1.freeze_panes(2, 0)
    #格式控制
    StylePortTitle = WorkBook.add_format({'bold': 1, 'align':'center','font_color':'blue'})
    StyleClassTitle = WorkBook.add_format({'bold': 1,'left':1})
    StyleGoodTitle = WorkBook.add_format({'left':1})
    StyleClassAmount = WorkBook.add_format({'bold': 1, 'num_format':'#,##0'})
    StyleClassRatio = WorkBook.add_format({'bold': 1, 'num_format':'0.0%'})
    StyleAmount = WorkBook.add_format({'num_format':'#,##0'})
    StyleRatio = WorkBook.add_format({'num_format':'0.0%'})

    for i in range(0, len(PortList)):
        port = PortList[i]
        #港口名称行
        #Sheet1.write(0,i*9,port)
        Sheet1.merge_range(0,i*9,0,i*9+8,port,StylePortTitle)
        totalamount = AmountInfo[0][port]
        totalsteel = AmountInfo[1][port]
        totaltrade = totalamount - totalsteel
        steeltotalratio, tradetotalratio = CalcRatio(totalamount, totalsteel, totaltrade)
        #标题行
        Sheet1.write(1, i*9, '',StyleGoodTitle)
        Sheet1.write(1, i*9+1, '总数量')
        Sheet1.write(1, i*9+2, '钢厂总数')
        Sheet1.write(1, i*9+3, '钢厂占比')
        Sheet1.write(1, i*9+4, '贸易商总数')
        Sheet1.write(1, i*9+5, '贸易商占比')
        Sheet1.write(1, i*9+6, '贸易商货权Top 1-3')
        Sheet1.write(1, i*9+7, '贸易商货权Top 4-6')
        Sheet1.write(1, i*9+8, '贸易商货权Top other')
        #合计数据行
        Sheet1.write(2, i*9, '合计',StyleClassTitle)
        Sheet1.write(2, i*9+1, totalamount,StyleClassAmount)
        Sheet1.write(2, i*9+2, totalsteel,StyleClassAmount)
        Sheet1.write(2, i*9+3, steeltotalratio,StyleClassRatio)
        Sheet1.write(2, i*9+4, totaltrade,StyleClassAmount)
        Sheet1.write(2, i*9+5, tradetotalratio,StyleClassRatio)
        ClassLineIndexDict = ClassLineIndex(GoodsClassName, GoodsClassList)
        for j in range(0,len(GoodsClassName)):
            classname = GoodsClassName[j]
            calsstotal = AmountInfo[2][port][classname]
            classsteel = AmountInfo[3][port][classname]
            classtrade = calsstotal - classsteel
            steelratio, traderatio = CalcRatio(calsstotal, classsteel, classtrade)
            Sheet1.write(ClassLineIndexDict[j]+3, i*9, classname,StyleClassTitle)
            Sheet1.write(ClassLineIndexDict[j]+3, i*9+1, calsstotal,StyleClassAmount)
            Sheet1.write(ClassLineIndexDict[j]+3, i*9+2, classsteel,StyleClassAmount)
            Sheet1.write(ClassLineIndexDict[j]+3, i*9+3, steelratio,StyleClassRatio)
            Sheet1.write(ClassLineIndexDict[j]+3, i*9+4, classtrade,StyleClassAmount)
            Sheet1.write(ClassLineIndexDict[j]+3, i*9+5, traderatio,StyleClassRatio)
            #if (j == 0) or (j == 1): #pass
            for k in range(0,len(GoodsClassList[j])):
                goodname = GoodsClassList[j][k]
                if goodname in AmountInfo[4][port].keys():
                    goodtotal = AmountInfo[4][port][goodname]
                    goodsteel = AmountInfo[5][port][goodname]
                    goodtrade = goodtotal - goodsteel
                    #goodship = ShipInfo[0][port][goodname]
                    #goodsteelship = ShipInfo[1][port][goodname]
                    goodothership = ShipInfo[2][port][goodname]
                    #GoodTop13, GoodTop46, GoodTopOther = TopShip(goodship)
                    #GoodSteelTop13, GoodSteelTop46, GoodSteelTopOther = TopShip(goodsteelship)
                    GoodOtherTop13, GoodOtherTop46, GoodOtherTopOther = TopShip(goodothership)
                else:
                    goodtotal = 0
                    goodsteel = 0
                    goodtrade = 0
                    #GoodTop13, GoodTop46, GoodTopOther = ('','','')
                    #GoodSteelTop13, GoodSteelTop46, GoodSteelTopOther = ('','','')
                    GoodOtherTop13, GoodOtherTop46, GoodOtherTopOther = ('','','')
                steelratio, traderatio = CalcRatio(goodtotal, goodsteel, goodtrade)
                Sheet1.write(ClassLineIndexDict[j]+4+k, i*9, goodname,StyleGoodTitle)
                Sheet1.write(ClassLineIndexDict[j]+4+k, i*9+1, goodtotal,StyleAmount)
                Sheet1.write(ClassLineIndexDict[j]+4+k, i*9+2, goodsteel,StyleAmount)
                Sheet1.write(ClassLineIndexDict[j]+4+k, i*9+3, steelratio,StyleRatio)
                Sheet1.write(ClassLineIndexDict[j]+4+k, i*9+4, goodtrade,StyleAmount)
                Sheet1.write(ClassLineIndexDict[j]+4+k, i*9+5, traderatio,StyleRatio)
                Sheet1.write(ClassLineIndexDict[j]+4+k, i*9+6, GoodOtherTop13)
                Sheet1.write(ClassLineIndexDict[j]+4+k, i*9+7, GoodOtherTop46)
                Sheet1.write(ClassLineIndexDict[j]+4+k, i*9+8, GoodOtherTopOther)
    WorkBook.close()

### 区分书写csv还是xlsx ###
def WriteSummary(Flag,ResFileName,AmountInfo,ShipInfo, PortList, GoodsClassName, GoodsClassList):
    if (Flag == 'txt') or (Flag == 'csv'):
        WriteTXT(ResFileName,AmountInfo, ShipInfo, PortList, GoodsClassName, GoodsClassList)
    elif Flag == 'xlsx':
        WriteExcel(ResFileName,AmountInfo, ShipInfo, PortList, GoodsClassName, GoodsClassList)
    elif ('txt' in Flag) and ('xlsx' in Flag):
        WriteTXT(ResFileName,AmountInfo, ShipInfo, PortList, GoodsClassName, GoodsClassList)
        WriteExcel(ResFileName,AmountInfo, ShipInfo, PortList, GoodsClassName, GoodsClassList)

### 读取港口数据中的港口顺序 ###
def PortOrder(PortsData):
    PortList = []
    TempList = []
    for Rec in PortsData:
        TempList.append(Rec[-1])
    for i in range(len(TempList)):
        if i == 0 :
            PortList.append(TempList[i])
        else:
            if TempList[i] not in PortList:
                PortList.append(TempList[i])
    return PortList

def WritePortsData(PortList,PortsData):
    filename = '_'.join(PortList)
    now = datetime.datetime.now().strftime('%Y%m%d')#_%H%M')
    filename = filename+'-'+now+'.xlsx'
    WorkBook = xlsxwriter.Workbook(filename)
    Sheet1 = WorkBook.add_worksheet()
    Sheet1.freeze_panes(1, 0)
    for i in range(len(PortsData)+1):
        Sheet1.set_row(i,16)
    # 格式控制
    StyleTitle = WorkBook.add_format({'bold': 1, 'align': 'center', 'font_color': 'blue'})
    StyleCenter = WorkBook.add_format({'align': 'center'})
    StyleAmount = WorkBook.add_format({'num_format': '#,##0.00'})
    # 标题行
    Sheet1.write(0, 0, '原行号',StyleTitle)
    Sheet1.write(0, 1, '港口名称',StyleTitle)
    Sheet1.write(0, 2, '货主',StyleTitle)
    Sheet1.write(0, 3, '品种',StyleTitle)
    Sheet1.write(0, 4, '数量',StyleTitle)
    Sheet1.write(0, 5, '日期',StyleTitle)
    Sheet1.set_column(0, 0, 8)
    Sheet1.set_column(1, 1, 12)
    Sheet1.set_column(2, 3, 15)
    Sheet1.set_column(4, 5, 12)
    # 数据行
    for i in range(len(PortsData)):
        Rec = PortsData[i]
        Sheet1.write(i+1,0,Rec[0],StyleCenter)
        Sheet1.write(i+1,1,Rec[5],StyleCenter)
        Sheet1.write(i+1,2,Rec[1])
        Sheet1.write(i+1,3,Rec[2])
        Sheet1.write(i+1,4,Rec[3],StyleAmount)
        Sheet1.write(i+1,5,Rec[4],StyleCenter)
    WorkBook.close()
    print('汇总数据写入"%s"文件中.' % filename)