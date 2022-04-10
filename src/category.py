class Category:
    def __init__(self, categoryNo, rate):
        self.categoryNo = categoryNo
        self.rate = float(rate)
        self.quantity = 0
        self.cost = 0
        self.componentList = []

    def calculate_cost(self):
        self.cost = self.quantity * self.rate

    def add_component(self, component):
        self.componentList.append(component)
        if component.area == 0:
            self.quantity += component.surfaceArea
        else:
            self.quantity += component.area

    def clear_category(self):
        self.quantity = 0
        self.cost = 0
        self.componentList.clear()