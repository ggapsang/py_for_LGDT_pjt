import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles import colors
from openpyxl.styles import Color
import random
import time
import datetime

def genMDMid(ldWs) :
    print("STEP 1st : 해더값 세팅 중...")

    def makingMDMid (ldWs, mdmIdLst, mdmIdLst_rowNo) :
        print("MDM ID 생성 중...")
        rowNo = 2
        v = 1
        for row in ldWs.iter_rows(min_row=2) :
            print(v,"의 mdm id 생성")
            if row[mdmUp_col].value != "X" :
                naming_codes = []
                for cell_col in ldHdrNoDict.values() :
                    if cell_col != mdmId_col and cell_col != mdmUp_col and row[cell_col].value is not None :
                        naming_codes.append(str(row[cell_col].value))
                row[mdmId_col].value = '-'.join(naming_codes) # row[key_col].value =mdmID
                mdmIdLst.append(row[mdmId_col].value)
                mdmIdLst_rowNo.append(rowNo)
            rowNo = rowNo + 1
            v=v+1

    def checkingDupl(mdmIdDict, indiv_dict) :
        print("중복 MDM ID 확인 중...")
        seen = []
        for ky, val in mdmIdDict.items() :
            if val not in seen :
                seen.append(val)
                indiv_dict[ky] = val

        return indiv_dict

    hdrLst = ["MDM 등록 여부", "MDM 설비 ID", "사용 유체 코드", "태그 시리얼 번호", "배관 사이즈", "배관 사양 코드", "배관 보온 코드", "배관 트레이싱 코드", "배관 자켓 코드"]
    key = "MDM 설비 ID" 
    key2 = "MDM 등록 여부"
    key3 = "태그 시리얼 번호"
    ldHdrNoDict = {}

    # print(ldWs[1])
    for i in ldWs[1] : # 1은 해더가 있는 행의 값
        if i.value in hdrLst and i.value not in ldHdrNoDict.keys() :
            ldHdrNoDict[i.value] = i.column - 1

    mdmId_col = ldHdrNoDict.get(key)
    mdmUp_col = ldHdrNoDict.get(key2)
    srNo_col = ldHdrNoDict.get(key3)


    print("2nd STEP : 1차 MDM ID 생성 중...")
    mdmIdLst = []
    mdmIdLst_rowNo = []
    makingMDMid(ldWs=ldWs, mdmIdLst=mdmIdLst, mdmIdLst_rowNo=mdmIdLst_rowNo)
    mdmIdDict = dict(zip(mdmIdLst_rowNo, mdmIdLst))


    print("3rd STEP : 중복 체크 리스트 생성 중...")
    checklst = []
    indiv_dict = {}
    checkingDupl(mdmIdDict, indiv_dict)
    checklst = [x for x in mdmIdDict.keys() if x not in indiv_dict.keys()]
    print("mdm id 중복 행 : ", checklst)
    print("mdm id 중복 개수: ", len(checklst))


    print("4th STEP : 중복 MDM ID 변경 중...")
    rowNo = 2
    while len(checklst) > 0 :
        for rowNo in checklst :
            for row in ldWs.iter_rows(min_row=rowNo, max_row=rowNo) :
                # 중복 mdm id를 갖는 값의 serial no 변경
                srNoSplit = list(str(row[srNo_col].value))
                newSrNo_lst = []
                for i in srNoSplit :
                    try :
                        i = int(i)
                        i = random.randrange(0, 10)
                        newSrNo_lst.append(str(i))
                    except :
                        newSrNo_lst.append(str(i))
                row[srNo_col].value = ''.join(newSrNo_lst)
                # print(row[srNo_col].value)

                # 변경된 serial no 칼라마킹
                ldWs.cell(rowNo, srNo_col+1).fill = PatternFill(fill_type='solid', fgColor=Color('FFFF00'))

                # 변경된 serial no로 새 mdm id 생성
                naming_codes = []
                for cell_col in ldHdrNoDict.values() :
                    if cell_col != mdmId_col and cell_col != mdmUp_col and row[cell_col].value is not None :
                        naming_codes.append(str(row[cell_col].value))
                row[mdmId_col].value = '-'.join(naming_codes) # row[key_col].value =mdmID
                # print(row[mdmId_col].value)

                # mdm id 딕셔너리에 새로 반영함
                mdmIdDict[rowNo] = row[mdmId_col].value

        # 중복 체크 루프
        checklst = []
        indiv_dict = {}
        checkingDupl(mdmIdDict, indiv_dict)
        checklst = [x for x in mdmIdDict.keys() if x not in indiv_dict.keys()]
        print(len(checklst))
    print("루프문 종료. 중복 태그 없음")

start = time.time()

print("1st STEP : 파일 불러오는 중..")
path = "C:/Users/dayoung.kweon/Documents/00_Python/"
loadFilename = input("파일이름 : ") + ".xlsx"
loadWb = openpyxl.load_workbook(path + loadFilename)
# loadWb.save(path + "백업파일/" + loadFilename)

for x in ["배관_BD2", "배관_BRU"] :
    ldWs = loadWb[x]
    genMDMid(ldWs)

print("저장 중...")
loadWb.save(path + loadFilename)

end = time.time()
sec = end - start
result = str(datetime.timedelta(seconds=sec)).split(".")
print(result[0], "완료")
