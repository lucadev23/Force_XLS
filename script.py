import io
import msoffcrypto
import openpyxl

def createWords(characterList, lenght):
    complete_list = []
    for current in range(lenght):
        a = [i for i in characterList]
        for y in range(current):
            a = [i+x for i in characterList for x in a]
        complete_list = complete_list+a
    return complete_list

def bruteForce(pathFile, words):
    decrypted_workbook = io.BytesIO()
    for i in words:
        try:
            with open('password.xls', 'rb') as file:
                office_file = msoffcrypto.OfficeFile(file)
                office_file.load_key(password=str(i))
                print("Trovato-> "+str(i))
                return 1
                #office_file.decrypt(decrypted_workbook)
        except:
            continue
    return 0

listaCaratteri = 'abcdefghijklmnopqrstuvwxyz0123456789'
path = "password.xls"

if __name__ == "__main__":
    lunghezza=1
    while(1):
        words = createWords(listaCaratteri, lunghezza)
        if bruteForce(path, words) == 1:
            break
        lunghezza+=1