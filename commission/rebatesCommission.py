import pandas as pd
from utils.functions import tax_calcs

payments = pd.read_csv("Commission+-+Payments+Received.csv", skiprows=6, parse_dates=["Date tax rebate received from HMRC"])
sales = pd.read_csv("Commission+-+Sales+Received.csv", skiprows=6)
booklets = pd.read_csv("Commission+-+Submitted+Booklets.csv", skiprows=6)

#https://portal.cirrusresponse.com/caco_g_callLog.asp?back=&displayType=showme&reportType=datafiles&sortOrder=countDesc&callsType=allcalls&ctOp=55781143&scopeType=vcc_operators%2B55781143&dateRangeType=s&sd=28&sm=aug&sy=2022&st=00%3A00&ed=1&em=oct&ey=2022&et=00%3A00&rdp=28-Aug-2022+-+30-Sep-2022&rdtp=28-Aug-2022+00%3A00+-+30-Sep-2022+23%3A59&tsF=00%3A00&tsT=00%3A00&tsWD=1&tsWD=2&tsWD=3&tsWD=4&tsWD=5&tsWD=6&tsWD=7&email=advancecontractingsolutions%40charterhouse.com&osIA=on&mtTS=10&mtTS=11&mtTS=20&mtTS=30&mtTS=40&mtTS=50&mtTF=&ctLTmt=sw&ctTS=&ctCLImt=sw&ctCLIS=&ctDNmt=sw&ctDDIS=&ctDNCLImt=sw&ctDNCLI=&ctNNmt=sw&ctNN=&ctABmt=ac&ctAB=0&ctMinD=0&ctMaxD=999999&ctMinOD=0&ctMaxOD=999999&ctDT=any&ctDmt=sw&ctSD=&ctST=a&ctDRG=&ctDRS=&ctSC=allcalls&ctFR=&ctHO=nofilter&ctPO=nofilter&ctSMSO=nofilter&ctVCCf=&ctVCCv=&ctQID=&ctCC=&ctCCd=&ctCNmt=sw&ctCN=&bbSBr=on&bbSBy=on&bbOptsP=y&oabps=qh&ocSCT=10&ugG=timeline&ugT=count&ugS=-1&dfOptsP=y&dfISH=on&dfIC=on&dfFL=&dfCDRVCC=&dfPE=on&dfV=2&rcTN=&rcWT=sd&raD=call&oaShown=y&oaSALT=y&oaSALTD=1&oaSc=0&oaICIDs=&oaIOIDs=&oaIIO=on&oaIS=on&oaICC4WS=on&luLL=0&luSW=60&QR8AS=0&QR8TA=60&QR8GB=d&qSBG=on&qBGO=MAX&qDH=on&b=h&cuT=3&wlIP=&lpEx=y&fbGT=0&icWS=1440&icOC=&oneCall=&specificNode=&ocSCT=10&opAT=20&opIO=both&opEO=ex&opGF=type&oOOt=nonopnonenq&oHTGB=caco_customers&ocQT=20&ocQT1=20&ocQT2=0&ocQT3=0&tqANO=0&tqABO=0&ocSQT=10&ocSAT=0&ocAT=0&tqGB=service&mtTN=-1&mtG=r&mtGL=nc&tT=20&tGB=caco_numbers&gM=calls&gTb=auto&hmSS=1h&tG=0&gT=5&gSS=0&mtGO=0&mtTGL=0&mtTGS=0&ibGB=c&gMT=connectminutes&tMGB=caco_numbers&gRTI=5&ctIn=all&mcLT=all&clSO=inboundStartTime+ASC&rDGS=dy&ScopeTypeChannelSearch=0&qID=0&qIDL=&bdRMID=0&assSTX=200&assSSID=-999&assOwn=-999&assCat=-999&assCrSi=-999&assCsCu=0&assEx=-999&assExCu=0&assSt=-999&assRR=-999&assHoTy=-999&assExCols=0&nobanner=

commission = pd.DataFrame([], index=pd.period_range(start = f'2022-08-28', end = f'2022-10-01', freq = 'D'))

print(payments["Date tax rebate received from HMRC"])

for i in commission.index:
    
    print(payments.loc[payments["Date tax rebate received from HMRC"] == i])

commission.to_csv("Rebates Commission.csv")