# coding=utf-8
from ReadFunc import *
from WriteFunc import *

### 需要用户定义的参数 ###
PortFiles = ['岚山现货表 08.14.xls','岚桥港现货表(2018.08.13).xls','曹妃甸实业进出存统计表18.8.12.xls','曹妃甸弘毅散货库存0813.xls','曹妃甸矿三货运部出入库日报-2018.8.13.xls','连云港港存贸易矿8.12.xls','青岛日报表20180815.xls','京唐港8月27日库存 .xls','日照8.20.xls']   #港口数据文件
#PortFiles = ['日照8.20.xls']   #港口数据文件
StdPort = ['曹妃甸实业','曹妃甸弘毅','曹妃甸矿三','京唐','岚桥','岚山','连云港','青岛','日照']

### 读取各港口信息 ###
PortsData = CollectPortData(PortFiles,StdPort)

### 汇总各港口信息，统一格式输出 ###
PortList = PortOrder(PortsData)
WritePortsData(PortList,PortsData)