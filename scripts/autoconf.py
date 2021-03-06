import os
import pip

try:
    import xlrd
except ImportError:
    pip.main(["install", "xlrd"])

wb = xlrd.open_workbook("wb.xls")
sheet = wb.sheet_by_index(0)

help = """<<<help>>>
help - pokazuje dostepne opcje
define - definiowanie hostów
services - definiowanie usług
ping - definiowanie usługi ping
exit - opuszcza program
"""
hosts = """<<<hosts>>>
help - pokazuje dostepne rodzaje hostów
windows - windows-server
camera - generic-camera
printer - generic-printer
back - opuszcza sekcję definicji hostów
exit - opuszcza program
"""
services = """<<<services>>>
help - pokazuje dostepne rodzaje usług
server - usługi serwerowe
printer - usługi drukarek
back - opuszcza sekcję definicji usług
exit - opuszcza program
"""
server_services = """<<<server services>>>
help - pokazuje dostepne rodzaje hostów
uptime - sprawdza jak dlugo serwer jest włączony
cpuload - sprawdza zużycie procesora
memusage - sprawdza zużycie pamięci
nsclient ver - sprawdza wersję programu nsclient++
drive space - sprawdza dostępne miejsce na dysku C:
updates - sprawdza dostępność aktualizacji systemu Windows
back - opuszcza sekcję definicji hostów
exit - opuszcza program
"""
printer_services = """<<<printer services>>>
help - pokazuje dostepne rodzaje hostów
windows - windows-server
camera - generic-camera
printer - generic-printer
back - opuszcza sekcję definicji hostów
exit - opuszcza program
"""

os.system("CLS")
check = input(
    """<<<UWAGA>>>
Program usunie wszystkie utworzone wcześniej pliki, sprawdź czy nie są już potrzebne.
Naciśnij ENTER żeby zatwierdzić.
"""
)
files = ["camera.txt", "windows.txt", "printer.txt", "ping.txt", "server_services.txt"]

if check in "":
    for file in files:
        try:
            os.remove(file)
        except:
            pass
    os.system("CLS")
else:
    os.exit()

option = None
passed = ["exit", "back", "help"]
while not (option == "exit"):
    print(help)
    option = input("Podaj nazwe templatu: ")
    os.system("CLS")

    if option == "define":
        options = ["windows", "camera", "printer"]
        while not (option == "exit" or option == "back"):
            print(hosts)
            option = input("Podaj rodzaj hostów do zdefiniowania: ")
            os.system("CLS")

            if option == "windows":
                use = "windows-server"

            elif option == "camera":
                use = "generic-camera"

            elif option == "printer":
                use = "generic-printer"

            else:
                if option not in passed:
                    print('\n"' + option + '" - taki rodzaj hostów nie istnieje\n')

            if option in options:
                with open(option + ".txt", "w") as list:
                    for pos in range(sheet.nrows):
                        ipaddress = sheet.cell_value(pos, 0)
                        hostname = sheet.cell_value(pos, 1)

                        list.write("define host{\n")
                        list.write("\tuse\t\t" + use + "\n")
                        list.write("\thost_name\t" + hostname + "\n")
                        list.write("\talias\t\t" + hostname + "\n")
                        list.write("\taddress\t\t" + ipaddress + "\n")
                        if option == "printer":
                            list.write("\thostgroups\tnetwork-printers\n")
                            list.write("\tparents\t\tTYC-PRINT-2\n")
                        if option == "camera":
                            list.write("\thostgroups\tIPCAM\n")
                        list.write("\t}\n")

                list.close()
                print("\nkonfiguracja zapisana do " + option + ".txt\n")

    elif option == "services":
        while not (option == "exit" or option == "back"):
            print(services)
            option = input("Podaj rodzaj usług do zdefiniowania: ")
            os.system("CLS")

            if option == "server":
                options = [
                    "uptime",
                    "cpuload",
                    "memusage",
                    "nsclientver",
                    "drivespace",
                    "updates",
                ]
                while not (option == "exit" or option == "back"):
                    print(server_services)
                    option = input(
                        "Podaj usługi do zdefiniowania oddzielone spacją: "
                    ).split()
                    os.system("CLS")
                    os.remove("server_services.txt")
                    for option in option:
                        if option == "uptime":
                            service = "Uptime"
                            command = "check_nt!UPTIME"

                        elif option == "cpuload":
                            service = "CPU Load"
                            command = "check_nt!CPULOAD!-l 5,80,90"

                        elif option == "memusage":
                            service = "Memory Usage"
                            command = "check_nt!MEMUSE!-w 90 -c 95"

                        elif option == "nsclient ver":
                            service = "NSClient++ Version"
                            command = "check_nt!CLIENTVERSION"

                        elif option == "drive space":
                            service = "Drive Space C:\ System"
                            command = "check_nt!USEDDISKSPACE!-l C -w 80 -c 90"

                        elif option == "updates":
                            service = "Windows Updates"
                            command = "check_nrpe!check_updates"

                        else:
                            if option not in passed:
                                print('\n"' + option + '" - taka usługa nie istnieje\n')

                        if option in options:
                            with open("server_services.txt", "a+") as list:
                                for pos in range(sheet.nrows):
                                    hostname = sheet.cell_value(pos, 1)

                                    list.write("define service{\n")
                                    list.write("\tuse\t\t\tgeneric-service\n")
                                    list.write("\thost_name\t\t" + hostname + "\n")
                                    list.write(
                                        "\tservice_description\t" + service + "\n"
                                    )
                                    list.write("\tcheck_command\t\t" + command + "\n")
                                    list.write("\t}\n")

                    list.close()
                    print("\nkonfiguracja zapisana do server_services.txt\n")

            elif option == "printer":
                options = [
                    "status",
                    "display",
                    "model",
                    "pagecount",
                    "maintenance",
                    "imagingunit",
                    "wastetoner",
                    "blackcartridge",
                    "cyantoner",
                    "magentatoner",
                    "yellowtoner",
                ]
                while not (option == "exit" or option == "back"):
                    print(printer_services)
                    option = input(
                        "Podaj usługi do zdefiniowania oddzielone spacją: "
                    ).split()
                    os.system("CLS")
                    os.remove("printer_services.txt")
                    for option in option:
                        if option == "uptime":
                            service = "Uptime"
                            command = "check_nt!UPTIME"

                        elif option == "cpuload":
                            service = "CPU Load"
                            command = "check_nt!CPULOAD!-l 5,80,90"

                        elif option == "memusage":
                            service = "Memory Usage"
                            command = "check_nt!MEMUSE!-w 90 -c 95"

                        elif option == "nsclient ver":
                            service = "NSClient++ Version"
                            command = "check_nt!CLIENTVERSION"

                        elif option == "drive space":
                            service = "Drive Space C:\ System"
                            command = "check_nt!USEDDISKSPACE!-l C -w 80 -c 90"

                        elif option == "updates":
                            service = "Windows Updates"
                            command = "check_nrpe!check_updates"

                        else:
                            if option not in passed:
                                print('\n"' + option + '" - taka usługa nie istnieje\n')

                        if option in options:
                            with open("printer_services.txt", "a+") as list:
                                for pos in range(sheet.nrows):
                                    hostname = sheet.cell_value(pos, 1)

                                    list.write("define service{\n")
                                    list.write("\tuse\t\t\tgeneric-service\n")
                                    list.write("\thost_name\t\t" + hostname + "\n")
                                    list.write(
                                        "\tservice_description\t" + service + "\n"
                                    )
                                    list.write("\tcheck_command\t\t" + command + "\n")
                                    list.write("\t}\n")

                    list.close()
                    print("\nkonfiguracja zapisana do printer_services.txt\n")

            else:
                if option not in passed:
                    print('\n"' + option + '" - taki rodzaj usług nie istnieje\n')

    elif option == "ping":
        with open(option + ".txt", "w") as list:
            for pos in range(sheet.nrows):
                hostname = sheet.cell_value(pos, 1)

                list.write("define service{\n")
                list.write("\tuse\t\t\tgeneric-service\n")
                list.write("\thost_name\t\t" + hostname + "\n")
                list.write("\tservice_description\tPING\n")
                list.write("\tcheck_command\t\tcheck_ping!200.0,20%!400.0,60%\n")
                list.write("\tcheck_interval\t\t5\n")
                list.write("\tretry_interval\t\t1\n")
                list.write("\t}\n")

            list.close()
            print("\nkonfiguracja zapisana do " + option + ".txt\n")

    else:
        if option not in passed:
            print('\n"' + option + '" - taki template nie istnieje\n')
