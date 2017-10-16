import xlsxwriter

class vehicle:
    def __init__(self, id, licence_plate, brand, v_type, VIN, mileage, last_maintenance):
        self.id = id
        self.licence_plate = licence_plate
        self.brand = brand
        self.v_type = v_type
        self.VIN = VIN
        self.mileage = mileage
        self.last_maintenance = last_maintenance

    def get_vehicle(self):
        return self.brand + " " + self.v_type

def all_inputs(vehicles):
    for index, vehicle in enumerate(vehicles):
        print "ID: " + str(vehicle.id)  # index is an order number of the contact object in the contacts list
        print vehicle.get_vehicle()
        print vehicle.licence_plate
        print vehicle.VIN
        print vehicle.mileage
        print vehicle.last_maintenance
    if not vehicles:
        print "database is empty"

def new_input(vehicles):
    licence_plate = raw_input("Licence plate: ")
    brand = raw_input("Brand: ")
    v_type = raw_input("Type: ")
    VIN = raw_input("VIN: ")
    mileage = raw_input("Mileage: ")
    last_maintenance = raw_input("Last date of the maintenance:")
    id = len(vehicles) + 1

    new = vehicle(id=id, licence_plate=licence_plate, brand=brand, v_type=v_type, VIN=VIN, mileage=mileage, last_maintenance=last_maintenance)
    vehicles.append(new)
    print new.get_vehicle() + " successfully added."

def edit_input(vehicles):
    print "ID of input you'd like to edit:"
    for index, vehicle in enumerate(vehicles):
        print str(vehicle.id) + ") " + vehicle.get_vehicle()
    selected_id = int(raw_input("Select ID: "))
    selected_vehicle = None

    for index, vehicle in enumerate(vehicles):
        if selected_id == vehicle.id:
            selected_vehicle = vehicle
            break

    if selected_vehicle is not None:
        licence_plate = raw_input("Licence plate (" + selected_vehicle.licence_plate + "): ")
        mileage = raw_input("Mileage (" + selected_vehicle.mileage + "): ")
        last_maintenance = raw_input("Last date of the maintenance (" + selected_vehicle.last_maintenance + "): ")
        if licence_plate != '':
            selected_vehicle.licence_plate = licence_plate
        if mileage != '':
            selected_vehicle.mileage = mileage
        if last_maintenance != '':
            selected_vehicle.mileage = last_maintenance

def delete_input(vehicles):
    print "ID of input you'd like to delite:"
    for index, vehicle in enumerate(vehicles):
        print str(index) + ") " + vehicle.get_vehicle()
    selected_id = raw_input("Select ID: ")
    selected_vehicle = vehicles[int(selected_id)]
    vehicles.remove(selected_vehicle)
    print "successfully removed"

def export_database_to_excel(vehicles):
    workbook = xlsxwriter.Workbook('database.xlsx')
    worksheet = workbook.add_worksheet()

    for (index, label) in enumerate(['ID', 'licence plate', 'brand', 'type', 'VIN', 'mileage', 'last maintenance']):
        worksheet.write(0, index, label)

    row = 1

    for index, vehicle in enumerate(vehicles):
        worksheet.write(row, 0, vehicle.id)
        worksheet.write(row, 1, vehicle.licence_plate)
        worksheet.write(row, 2, vehicle.brand)
        worksheet.write(row, 3, vehicle.v_type)
        worksheet.write(row, 4, vehicle.VIN)
        worksheet.write(row, 5, vehicle.mileage)
        worksheet.write(row, 6, vehicle.last_maintenance)
        row += 1

    workbook.close()

def main():
    print "Vehicle database!\n"
    A = vehicle(id=1, licence_plate="MBxxxxx", brand="Mercedez", v_type="C-class", VIN="WZZZXYF123456", mileage="12345", last_maintenance="12.7.2017")
    B = vehicle(id=2, licence_plate="MB1xxxx", brand="Renault", v_type="Traffic", VIN="WZZZXYdf123456", mileage="167890", last_maintenance="6.3.2017")
    C = vehicle(id=3, licence_plate="MB1xxxx", brand="Renault", v_type="Traffic", VIN="WZZZXYdf123456", mileage="167890", last_maintenance="6.3.2017")
    vehicles = [A, B, C]

    while True:
        print "\na) See all your contacts"
        print "b) Add a new contact"
        print "c) Edit a contact"
        print "d) Delete a contact"
        print "e) Export database to excel"
        print "f) Quit the program."
        selection = raw_input("\nHow would you like to proceed? (a, b, c, d or e): ")

        if selection.lower() == "a":
            all_inputs(vehicles)
        elif selection.lower() == "b":
            new_input(vehicles)
        elif selection.lower() == "c":
            edit_input(vehicles)
        elif selection.lower() == "d":
            delete_input(vehicles)
        elif selection.lower() == "e":
            export_database_to_excel(vehicles)
        elif selection.lower() == "f":
            print "Goodbye!"
            break
        else:
            print "Invalid selection\n\n"
            continue

if __name__ == "__main__":  # this means that if somebody ran this Python file, execute only the code below
    main()