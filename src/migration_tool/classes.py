def check_formula(cell):
    """Check whether a cell contains a formula"""
    if cell.data_type == "f":
        return True
    else:
        return False


class shapes(object):
    """
    Class for shapes (ex left:1, top:1, right:3, bottom:3).
    Initialized with a tuple (or None if no shape is found).
    Provides methods to get the left, top, right, bottom, number of rows and columns of the shape,
    check if it is one cell, and
    build value from the shapes (list of lists).
    """

    def __init__(self, tuple):
        if tuple is None:
            self.shape = None
        else:
            self.shape = tuple

    # Getters
    def left(self):
        if self.shape is not None:
            return self.shape[0]

    def top(self):
        if self.shape is not None:
            return self.shape[1]

    def right(self):
        if self.shape is not None:
            return self.shape[2]

    def bottom(self):
        if self.shape is not None:
            return self.shape[3]

    def rows(self):
        if self.shape is not None:
            return self.bottom() - self.top() + 1

    def cols(self):
        if self.shape is not None:
            return self.right() - self.left() + 1

    def isonecell(self):
        if self.rows() == 1 and self.cols() == 1:
            return True
        else:
            return False

    def build_values(self, sheet):
        if self.shape is None:
            return None

        topleft = sheet.cell(row=self.top(), column=self.left())

        if self.isonecell():
            if check_formula(topleft):
                return None
            else:
                return [[topleft.value]]
        else:
            list_values = []
            found_formula = False

            for row in range(self.top(), self.bottom() + 1):
                row_values = []
                for col in range(self.left(), self.right() + 1):
                    cell = sheet.cell(row=row, column=col)
                    if check_formula(cell):
                        row_values = None
                        found_formula = True
                        break
                    else:
                        row_values.append(cell.value)
                if found_formula:
                    list_values = None
                    break
                list_values.append(row_values)

            return list_values


class values(object):
    """
    Class for values (ex [[1,2],[3,4]]).
    Initialized with a list of lists (or None if no values are found).
    Provides methods to get the values, the top left value, the first element of each row,
    to enlarge the values to fit an input shape,
    to add checkboxes (change None to False), and
    to paste the values into a sheet (iterating over rows and columns).
    """

    def __init__(self, list_of_lists):
        if list_of_lists is None:
            self.val = None
        else:
            self.val = list_of_lists
            self.height = lambda: len(self.val)
            self.width = lambda: len(self.val[0])

    # Getter
    def values(self):
        return self.val

    def topleft(self):
        if self.values() is not None:
            return self.values()[0][0]
        else:
            return None

    def first_element_row(self, double_list=True):
        if double_list:
            for row in range(self.height()):
                self.values()[row] = [self.values()[row][0]]
        else:
            list = []
            for row in range(self.height()):
                list.append(self.values()[row][0])
            return list

    def enlarged_range_correction(self, shape):
        if self.height() != shape.rows():
            for i in range(shape.rows() - self.height()):
                self.values().append([None])

    def add_checkboxes(self):
        for row in range(self.height()):
            if self.values()[row][0] is None:
                self.values()[row][0] = False

    def convert_month_to_numbers(self):
        month_dict = {
            "January": 1,
            "February": 2,
            "March": 3,
            "April": 4,
            "May": 5,
            "June": 6,
            "July": 7,
            "August": 8,
            "September": 9,
            "October": 10,
            "November": 11,
            "December": 12,
        }
        if self.values() is not None:
            if self.values()[1][0] in month_dict:
                self.values()[1][0] = month_dict[self.values()[1][0]]
            else:
                self.values()[1][0] = None

    def count_uniques(self):
        list = self.values()
        return len(set([item for sublist in list for item in sublist]))

    def paste(self, sheet, shape):
        if self.width() == 1:
            for row in range(shape.top(), shape.bottom() + 1):
                sheet.cell(row=row, column=shape.left()).value = self.values()[
                    row - shape.top()
                ][0]
        else:
            for row in range(shape.top(), shape.bottom() + 1):
                for col in range(shape.left(), shape.right() + 1):
                    sheet.cell(row=row, column=col).value = self.values()[
                        row - shape.top()
                    ][col - shape.left()]
