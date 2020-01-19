from openpyxl import load_workbook


class LifeExpectancyReader:

    def __init__(self):

        self.workbook = load_workbook("life_expectancy_2014.xlsx", read_only=True)
        self.worksheet = self.workbook.active
        self.life_exp_dictionary = {}
        self.read_in_data(self.life_exp_dictionary)

    def read_in_data(self, life_exp_dic):
        for key in self.worksheet.values:  # .values brings us from cell  references to actual values. Processed in rows
            if isinstance(key[0], int):
                print(key)
                life_exp_dic[key[0]] = {"Country": key[1],
                                        "Life Expectancy": key[2]}

        # self.life_exp_dictionary['Chad'] = 4
        # self.life_exp_dictionary['Chad'] = [223, 4]

    def print_top_ten_formatted(self):
        print("Top Life Expectancies (Av.)")
        print("|    ID      |   Country                 |   Life Expectancy")
        print("|____________|___________________________|___________________")
        for key in self.life_exp_dictionary:
            if key in range(1, 11):
                print("|", str(key), " "*(9-len(str(key))), "|",
                      self.life_exp_dictionary[key]["Country"], ' '*(24-len(self.life_exp_dictionary[key]["Country"])),
                      "|",
                      self.life_exp_dictionary[key]["Life Expectancy"],)
        print("\n")

    def print_bottom_ten_formatted(self):
        print("Bottom Life Expectancies (Av.)")
        print("|    ID      |   Country                 |   Life Expectancy")
        print("|____________|___________________________|___________________")
        for key in range(len(self.life_exp_dictionary) - 10, len(self.life_exp_dictionary)):
            print("|", str(key), " " * (9 - len(str(key))), "|",
                  self.life_exp_dictionary[key]["Country"], ' ' * (24 - len(self.life_exp_dictionary[key]["Country"])),
                  "|",
                  self.life_exp_dictionary[key]["Life Expectancy"])
        print("\n")

    def print_average(self):
        total = 0
        counter = 0
        for life_ex_value in range(1, len(self.life_exp_dictionary)):
            total += self.life_exp_dictionary[life_ex_value]["Life Expectancy"]
            counter += 1
        print("Average Global Life Expectancy is " + "%.2f" % (total/counter) + " years old.\n")
      # print("{1:.2f} {0:.2f}.format(pi, pi/2))
    def print_rank_details(self, rank):
        for key in range(1, len(self.life_exp_dictionary)):
            if key == rank:
                print("You selected rank " + str(rank) + "\n" + "Details - Country:  "
                      + self.life_exp_dictionary[key]["Country"] + "   "
                      + "Life Expectancy:  " + str(self.life_exp_dictionary[key]["Life Expectancy"]))
                if (self.life_exp_dictionary[key]["Life Expectancy"]) < 60:
                    print("Don't move there.")




x = LifeExpectancyReader()

x.print_top_ten_formatted()
x.print_bottom_ten_formatted()
x.print_average()
x.print_rank_details(40)