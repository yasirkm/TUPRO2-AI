from openpyxl import load_workbook
from pathlib import Path

FOLDER_PATH = Path(__file__).parent
EXCEL_PATH = FOLDER_PATH/"bengkel.xlsx"

def main():
    # Load excel file
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Sheet1"]
    service = {"buruk": 0, "biasa": 0, "bagus": 0}
    price   = {"murah": 0, "sedang": 0,"mahal": 0}
    rows = ws.rows
    next(rows)

    bengkel_list = [Bengkel(id.value, service.value, price.value) for id, service, price in rows]
    
    for bengkel in bengkel_list:
        print(bengkel.fuzzification())

    wb.close()

class Bengkel():
    def __init__(self, id, service, price):
        self.id = id
        self.service = service
        self.price = price

    def fuzzification(self):
        fuzzy_service = {}
        fuzzy_price = {}
        #MEMBERSHIP FUNCTION (CRISP INPUT -> FUZZY INPUT)
        #SERVICE CATEGORY
        fuzzy_service["buruk"] = max(0, (50-self.service)/(50-1))
        fuzzy_service["biasa"] = max(0, (self.service-20)/(40-20)) if self.service < 40 else 1 if self.service <= 60 else max(0,(80-self.service)/(80-60)) if self.service < 80 else 0
        fuzzy_service["bagus"] = max(0, (self.service-50)/(100-50))
        #PRICE CATEGORY
        fuzzy_price["murah"] = max(0, (5-self.price)/(5-1))
        fuzzy_price["sedang"] = max(0, (self.price-2)/(4-2)) if self.price < 4 else 1 if self.price <= 6 else max(0,(9-self.price)/(9-6)) if self.price < 9 else 0
        fuzzy_price["mahal"] = max(0, (self.price-5)/(10-5))

        return fuzzy_service, fuzzy_price

    def inference(self):
        '''inference rules '''
        fuzzy_service, fuzzy_price = self.fuzzification()
        inference = 

    
    #def membership_function(self):

    #def service_partition(self):
        
    # def price_partition(self):
        
    def defuzzification(self):
        #defuzzification = (service_partition * price_partition) / (service_partition + price_partition)


if __name__ == "__main__":
    main()
