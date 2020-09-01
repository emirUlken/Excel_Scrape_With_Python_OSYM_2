import xlrd
import json
import requests
import PyPDF2
import re

path = "TYT_TABAN_PUAN.xls"
API_ENDPOINT_DEPARTMENT_BULK = "http://localhost:8080/department/bulk"
API_ENDPOINT_UNIVERSITY_BULK = "http://localhost:8080/university/bulk"
API_ENDPOINT_POINTTYPE_BULK = "http://localhost:8080/pointtype/bulk"
API_ENDPOINT_KOSUL_BULK = "http://localhost:8080/kosulaciklama/bulk"
headers = {'Content-type': 'application/json', 'Accept': 'text/plain'}

inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)

names = []
codes = []
ogrenimSuresiList = []
puanTuruList = []
kontenjanList = []
okulBirinciKontenjanList = []
basariSirasiList = []
tabanList = []
kosulList = []
isFullList = []
uniIndexArray = []
codeIndexArray = []
pointIndexArray = []
kosulIndexArr = []
numberIndexArr = []
jsonArray = "["
jsonArrayUni = "["
jsonArrayPoint = "["
jsonArrayKosul = "["
count = 0
tempCount = 0
uni = ""
tempUni = ""
tempCode = ""
tempPoint = ""
PDFURL = "kosul ve aciklamalar.pdf"
# This is the contact list of universities at the end of the excel file. We remove it for scraping data.
removeStr = r'ABANT İZZET BAYSAL ÜNİ.  http://www.ibu.edu.trADNAN MENDERES ÜNİ. http://www.adu.edu.trAFYON KOCATEPE ÜNİ.  http://www.aku.edu.trAKDENİZ ÜNİVERSİTESİ http://www.akdeniz.edu.trANADOLU ÜNİVERSİTESİ http://www.anadolu.edu.trANKARA ÜNİVERSİTESİ http://www.ankara.edu.trATATÜRK ÜNİVERSİTESİ  http://www.atauni.edu.trBALIKESİR ÜNİ. http://www.balikesir.edu.tr BOĞAZİÇİ ÜNİVERSİTESİ http://www.boun.edu.trCELAL BAYAR ÜNİ.  http://www.bayar.edu.tr CUMHURİYET ÜNİ. http://www.cumhuriyet.edu.trÇANAKKALE ONSEKİZ MART ÜNİ.http://www.comu.edu.trÇUKUROVA ÜNİVERSİTESİ  http://www.cukurova.edu.trDİCLE ÜNİ. http://www.dicle.edu.trDOKUZ EYLÜL ÜNİVERSİTESİ http://www.deu.edu.trDUMLUPINAR ÜNİ. http://www.dumlupinar.edu.tr EGE ÜNİVERSİTESİ http://www.ege.edu.trERCİYES ÜNİVERSİTESİ http://www.erciyes.edu.trFIRAT ÜNİ.  http://www.firat.edu.trGALATASARAY ÜNİVERSİTESİ  http://www.gsu.edu.trGAZİ ÜNİVERSİTESİ  http://www.gazi.edu.trGAZİANTEP ÜNİVERSİTESİ  http://www.gantep.edu.tr GAZİOSMANPAŞA ÜNİ.  http://www.gop.edu.trGEBZE YÜKSEK TEKNOLOJİ ENS. http://www.gyte.edu.trHACETTEPE ÜNİVERSİTESİ http://www.hacettepe.edu.trHARRAN ÜNİ. http://www.harran.edu.trİNÖNÜ ÜNİ. http://www.inonu.edu.trİSTANBUL ÜNİVERSİTESİ  http://www.istanbul.edu.trİSTANBUL TEKNİK ÜNİVERSİTESİ   http://www.itu.edu.trİZMİR YÜKSEK TEKNOLOJİ ENS.  http://www.iyte.edu.tr KAFKAS ÜNİVERSİTESİ http://www.kafkas.edu.trK.MARAŞ SÜTÇÜ İMAM ÜNİ. http://www.ksu.edu.trKARADENİZ TEKNİK ÜNİVERSİTESİ http://www.ktu.edu.trKIRIKKALE ÜNİ. http://www.kku.edu.trKOCAELİ ÜNİ. http://www.kou.edu.trMARMARA ÜNİVERSİTESİ http://www.marmara.edu.trMERSİN ÜNİ. http://www.mersin.edu.trMİMAR SİNAN GÜZEL SAN.ÜNİ.http://www.msu.edu.trMUĞLA ÜNİ. http://www.mu.edu.trMUSTAFA KEMAL ÜNİ.  http://www.mku.edu.trNİĞDE ÜNİ. http://www.nigde.edu.trONDOKUZ MAYIS ÜNİVERSİTESİ  http://www.omu.edu.trORTA DOĞU TEKNİK ÜNİVERSİTESİ   http://www .odtu.edu.trOSMANGAZİ ÜNİVERSİTESİ  http://www.ogu.edu.tr PAMUKKALE ÜNİ.  http://www.pamukkale.edu.trSAKARYA ÜNİ.  http://www.sakarya.edu.trSELÇUK ÜNİVERSİTESİ  http://www.selcuk.edu.tr SÜLEYMAN DEMİREL ÜNİ. http://www.sdu.edu.trTRAKYA ÜNİVERSİTESİ  http://www.trakya.edu.trULUDAĞ ÜNİVERSİTESİ  http://www.uludag.edu.trYILDIZ TEKNİK ÜNİVERSİTESİ http://www.yildiz.edu.trYÜZÜNCÜ YIL ÜNİ. http://www.yyu.edu.trZONGULDAK KARAELMAS ÜNİ. http://www.karaelmas.edu.trGÜLHANE ASKERİ TIP AKADEMİSİ  http://www. gata.edu.trATILIM ÜNİVERSİTESİ  http://www.atilim.edu.trBAHÇEŞEHİR ÜNİVERSİTESİ http://www.bahcesehir.edu.trBAŞKENT ÜNİVERSİTESİ http://www.baskent.edu.trBEYKENT ÜNİVERSİTESİ http://www.beykent.edu.trBİLKENT ÜNİVERSİTESİ http://www.bilkent.edu.trÇAĞ ÜNİVERSİTESİ http://www.cag.edu.trÇANKAYA ÜNİVERSİTESİ  http://www.cankaya.edu.trDOĞUŞ ÜNİVERSİTESİ http://www.dogus.edu.trFATİH ÜNİVERSİTESİ  http://www.fatih.edu.trHALİÇ ÜNİVERSİTESİ http://www.halic.edu.trIŞIK ÜNİVERSİTESİ http://www.isik.edu.trİSTANBUL BİLGİ ÜNİVERSİTESİ  http://www. ibun.edu.trİSTANBUL KÜLTÜR ÜNİVERSİTESİ   http://www .iku.edu.trİSTANBUL TİCARET ÜNİ.  http://www .iticu .edu.trİZMİR EKONOMİ ÜNİ.  http://www.izmirekonomi.edu.trKADİR HAS ÜNİVERSİTESİ http://www.khas.edu.trKOÇ ÜNİVERSİTESİ http://www.ku.edu.trMALTEPE ÜNİ.  http://www.maltepe.edu.trOKAN ÜNİVERSİTESİ http://www.okan.edu.tr SABANCI ÜNİVERSİTESİ http://www.sabanciuniv.edu.tr TOBB EKONOMİ VE TEKNOLOJİ ÜNİ.http://www.etu.edu.trUFUK ÜNİ. http://www.ufuk.edu.trYAŞAR ÜNİVERSİTESİ  http://www.yasar.edu.tr YEDİTEPE ÜNİVERSİTESİ http://www.yeditepe.edu.trANADOLU KÜLTÜR VE EĞİTİM VAKFI  http://www.anadolubil.edu.trMERSİN İLAĞA EĞİTİM VE KÜLTÜR VAKFI  http://www.medet.edu.trDOĞU AKDENİZ ÜNİVERSİTESİ  http://www.emu.edu.trGİRNE AMERİKAN ÜNİVERSİTESİ  http://www.gau.edu.trLEFKE AVRUPA ÜNİVERSİTESİ  http://www.lefke.edu.trULUSLARARASI KIBRIS ÜNİ.http://www.ciu.edu.trYAKIN DOĞU ÜNİVERSİTESİ  http://www.neu.edu.trİNTERNET ADRESLERİ '


# Function for scraping koşul açıklamalar
def createKosul():
    kosulHashMap = {}
    completeString = ""

    pdfFileObj = open(PDFURL, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj, strict=False)

    for page in range(0, pdfReader.numPages):
        pageObj = pdfReader.getPage(page)
        completeString += pageObj.extractText().replace("ý", "ı").replace("þ", "ş").replace("ð",
                                                                                            "ğ").replace(
            "Ð", "Ğ").replace("Ý", "İ").replace("Þ", "Ş").replace("™", "\'").replace("ﬁ", "\"").replace("ﬂ",
                                                                                                        "\"").replace(
            "\n", " ")

    pageObj = pdfReader.getPage(0)
    completeString = re.sub('[0-9]+\(Adayların, koşul ve açıklamalar içerisinde yer alan burs koşullarının geçerliliği '
                            'hakkında ilgiliüniversiteden bilgi almaları yararlarına olacaktır.\)', '', completeString)

    completeString = re.sub('TABLO-3A, TABLO-3B  VE  TABLO-4\'TE  YER  ALAN YÜKSEKÖĞRETİM  PROGRAMLARININ  KOŞUL  VE  '
                            'AÇIKLAMALARI ', '', completeString)

    completeString = re.sub(removeStr, '', completeString)

    tempList = re.split("(Bk.[0-9]+\.)", completeString)

    tempList.pop(0)

    for s in tempList:
        if tempList.index(s) % 2 == 0:
            numberIndexArr.append(re.findall(r'\d+', s)[0])
        else:
            kosulIndexArr.append(s)

    i = 0
    for index in numberIndexArr:
        kosulHashMap[index] = kosulIndexArr[i]
        i += 1

    return kosulHashMap


# Our hashmap (dict) of koşul açıklamalar
kosulFinalMap = createKosul()

tempTempCount = 0
# Traverse the rows
for row in range(inputWorksheet.nrows):
    kosulTempArr = []
    row0 = inputWorksheet.cell_value(row, 0)
    row1 = inputWorksheet.cell_value(row, 1)
    # Üniversite ata
    if "ÜNİVERSİTESİ" in str(row1).strip() or "MESLEK" in str(row1).strip():
        uni = str(row1).strip().split(" (")[0]
        uni = uni.strip()
        if tempUni != uni:
            tempUni = uni
            uniIndexArray.append(uni)
        # uni = str(row1).strip().split(" (")[0]

    # Row fake ise atla
    if len(row1) == 0 or len(row0) == 0 or count == 1:
        count += 1
        continue

    row2 = inputWorksheet.cell_value(row, 2)
    row3 = inputWorksheet.cell_value(row, 3)
    row4 = inputWorksheet.cell_value(row, 4)
    row5 = inputWorksheet.cell_value(row, 5)
    row6 = inputWorksheet.cell_value(row, 6)
    row7 = inputWorksheet.cell_value(row, 7)
    row8 = inputWorksheet.cell_value(row, 8)

    code = str(row0).strip()

    if tempCode != code[0:4]:
        tempCode = code[0:4]
        codeIndexArray.append(code[0:4])
        # Create JSON Object for a single university
        y = {
            "name": "%s" % tempUni,
            "code": "%s" % tempCode
        }
        jsonArrayUni += json.dumps(y, ensure_ascii=False) + ","

    codes.append(code)

    name = str(row1).strip()
    names.append(name)

    ogrenimSuresi = str(row2).strip().split(".")[0]
    ogrenimSuresiList.append(ogrenimSuresi)

    puanTuru = str(row3).strip()

    # Create PointType array
    if tempPoint != puanTuru:
        tempPoint = puanTuru
        pointIndexArray.append(puanTuru)
        # Create JSON Object for a single point type object
        h = {
            "type": "%s" % tempPoint,
        }
        jsonArrayPoint += json.dumps(h, ensure_ascii=False) + ","

    puanTuruList.append(puanTuru)

    kontenjan = str(row4).strip().split(".")[0]
    kontenjanList.append(kontenjan)

    # Okul birinciliği kontenjanı kontrol et
    if len(str(row5)) == 0:
        okulBirinciKontenjan = str(0)
        okulBirinciKontenjanList.append(okulBirinciKontenjan)
    else:
        okulBirinciKontenjan = str(row5).strip().split(".")[0]
        okulBirinciKontenjanList.append(okulBirinciKontenjan)

    # Koşul kontrol et
    if len(str(row6)) == 0:
        kosulAciklama = []
        kosulTempArr = []
    else:
        kosulAciklama = str(row6).strip()
        for num in re.findall(r'\d+', kosulAciklama):
            if num not in numberIndexArr:
                kosulTempArr = []
                break
            else:
                kosulDict = {"id": "%s" % str(numberIndexArr.index(num) + 1)}
                kosulTempArr.append(kosulDict)
    # Kontenjan doldu mu kontrol et
    if len(str(row7)) == 0 or str(row7).strip() == "..." or str(row7).strip() == "----":
        basariSirasi = "0"
        taban = "0.0"
        isFull = False

        basariSirasiList.append(basariSirasi)
        tabanList.append(taban)
        isFullList.append(isFull)
    else:
        basariSirasi = str(row7).strip().split(".")[0]
        taban = str(row8).strip()
        isFull = True

        basariSirasiList.append(basariSirasi)
        tabanList.append(taban)
        isFullList.append(isFull)

    # Create JSON Object for department
    x = {
        "code": "%s" % code,
        "name": "%s" % name,
        "uni": {"id": uniIndexArray.index(uni) + 1},
        "depDetails": {
            "ogrenimSuresi": "%s" % ogrenimSuresi,
            "puanTuru": {"id": pointIndexArray.index(puanTuru) + 1},
            "kontenjan": kontenjan,
            "okulBirinciKontenjan": okulBirinciKontenjan,
            "kosulAciklama": kosulTempArr,
            "basariSirasi": basariSirasi,
            "taban": taban,
            "fullStatus": isFull
        }
    }

    #jsonArray += json.dumps(x, ensure_ascii=False) + ","
    # This is for limiting the amount of department to send to the server.
    # Currently it's limited do 1000 for testing purposes. You can remove this
    # by applying the code above instead of this if statement below.
    if tempTempCount < 1000:
        jsonArray += json.dumps(x, ensure_ascii=False) + ","
        tempTempCount += 1

    # Add JSONObject to JSONArray
    count += 1

# Create jsonArrayKosul
for key, value in kosulFinalMap.items():
    tempDict = {"number": "%s" % str(key), "description": "%s" % value}
    jsonArrayKosul += json.dumps(tempDict, ensure_ascii=False) + ","

jsonArray = jsonArray[:-1]
jsonArrayUni = jsonArrayUni[:-1]
jsonArrayPoint = jsonArrayPoint[:-1]
jsonArrayKosul = jsonArrayKosul[:-1]
jsonArray += "]"
jsonArrayUni += "]"
jsonArrayPoint += "]"
jsonArrayKosul += "]"
# Send POST request to the server
print("Sunucudan yanıt bekleniyor...")

# Send pointtype list
r = requests.post(url=API_ENDPOINT_POINTTYPE_BULK, data=str(jsonArrayPoint).encode("utf8"), headers=headers)
response = r.text
print("The response from PointType API is: %s" % response)

# Send kosul list
r = requests.post(url=API_ENDPOINT_KOSUL_BULK, data=str(jsonArrayKosul).encode("utf8"), headers=headers)
response = r.text
print("The response from Kosul API is: %s" % response)


# Send uni list
r = requests.post(url=API_ENDPOINT_UNIVERSITY_BULK, data=str(jsonArrayUni).encode("utf8"), headers=headers)
response = r.text
print("The response from Uni API is: %s" % response)

# Send dep list
r = requests.post(url=API_ENDPOINT_DEPARTMENT_BULK, data=str(jsonArray).encode("utf8"), headers=headers)
response = r.text
print("The response from Department API is: %s" % response)

print("Count = %s" % count)
