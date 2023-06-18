import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles import colors
from openpyxl.styles import Color
import random
import time
import datetime

def getHdrNoDict(loadWs, hdrlst, HdrNoDict) :
    for i in loadWs[1] : #해더가 시작되는 열
        if i.value in hdrlst and i.value not in HdrNoDict.keys() :
            HdrNoDict[i.value] = i.column - 1
    revHdrNodict = dict(map(reversed, HdrNoDict.items()))
    return revHdrNodict

def getPmsTable(loadWs, ldHdrNo, Table) :
    for i in loadWs.iter_rows(min_row= 2) :
        dic = {}
        for j in ldHdrNo :
            if j != "MAT+SIZE" :
                dic[j] = i[ldHdrNo[j]].value    
        Table[i[ldHdrNo.get("MAT+SIZE")].value] = dic

start = time.time()

print("1st STEP : 파일 불러오는 중..")
path = "C:/Users/dayoung.kweon/Documents/00_Python/"
loadFilename = "대산_POE_모수정리_230425_v0.8.xlsx"
loadWb = openpyxl.load_workbook(path + loadFilename)
# loadWb.save(path + "백업파일/" + loadFilename)

pmsWs = loadWb["배관 PMS"]
pmsWsHdr = ["MATERIAL CODE", "MAT+SIZE", "PLANT", "출처명", "출처명2", "출처명3", "타입 이름", "SIZE SCOPE(MIN)", "SIZE SCOPE(MAX)", "플랜지 래이팅", "플랜지 접촉면 타입", "배관 재질", "가스켓 재질", "배관 스케줄", "배관 두께"]
pmsWsHdrNo = {}

for i in ["배관_POE", "배관_LLDPE"] :
    pipWs = loadWb[i]
    # pipWs = loadWb["배관"] # 시트 이름_입력 수정 값
    pipWsHdr = ["MDM 등록 여부", "타입 이름", "배관 사양 코드", "사이즈2", "플랜지 래이팅", "플랜지 접촉면 타입", "배관 재질", "가스켓 재질", "배관 스케줄", "배관 두께"]
    pipWsValue = ["타입 이름", "플랜지 래이팅", "플랜지 접촉면 타입", "배관 재질", "가스켓 재질", "배관 스케줄", "배관 두께"]
    pipWsHdrNo = {}

    print("2nd STEP : PMS시트 해더 리스트 생성 중..")
    getHdrNoDict(loadWs=pmsWs, hdrlst=pmsWsHdr, HdrNoDict=pmsWsHdrNo)
    getHdrNoDict(loadWs=pipWs, hdrlst=pipWsHdr, HdrNoDict=pipWsHdrNo)
    print(pmsWsHdrNo)
    print(pipWsHdrNo)

    pmsTable = {}
    getPmsTable(loadWs=pmsWs, ldHdrNo=pmsWsHdrNo, Table=pmsTable)
    # print(pmsTable)

    for i in pipWs.iter_rows(min_row=2) :
        find_count = 0
        not_find_count = 0
        if i[pipWsHdrNo["MDM 등록 여부"]].value == "△" :
            for j in pmsTable :
                try :
                    if i[pipWsHdrNo["배관 사양 코드"]].value == pmsTable[j].get("MATERIAL CODE") and pmsTable[j].get("SIZE SCOPE(MIN)") <= i[pipWsHdrNo["사이즈2"]].value and i[pipWsHdrNo["사이즈2"]].value <= pmsTable[j].get("SIZE SCOPE(MAX)") :
                        find_count = find_count+1
                        for k in pipWsValue :
                            i[pipWsHdrNo[k]].value = pmsTable[j].get(k)
                except :
                    not_find_count = not_find_count + 1
                    continue
        else : 
            continue            
        print("찾은속성값 : ", find_count," 못찾은 속성값 : ", not_find_count)

    # print(pmsTable)

print("저장 중...")
loadWb.save(path + loadFilename)



end = time.time()
sec = end - start
result = str(datetime.timedelta(seconds=sec)).split(".")
print(result[0], "완료")
