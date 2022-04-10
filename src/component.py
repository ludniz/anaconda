class Component:
    def __init__(self, size, area, surfaceArea):
        self.size = size
        self.area = float(self.strip_unit(area))
        self.surfaceArea = float(self.strip_unit(surfaceArea))
        self.width = 0
        self.height = 0
        self.category = 0

    def set_width_and_height(self):
        if self.size.count('-') == 0:
            self.width = int(self.size.split('x')[0])
            self.height = int(self.size.split('x')[1])
        else:
            sizes = self.size.split('-')
            widths = []
            heights = []
            for item in sizes:
                widths.append(int(item.split('x')[0]))
                heights.append(int(item.split('x')[1]))
            self.width = min(widths)
            self.height = min(heights)

    def set_duct_category(self):
        if max(self.width, self.height) < 750 and (self.width + self.height) <= 1150:
            self.category = '1'
        elif max(self.width, self.height) < 750 and (self.width + self.height) > 1150:
            self.category = '2'
        elif 750 <= max(self.width, self.height) < 1350:
            self.category = '3'
        elif 1350 <= max(self.width, self.height) < 2100:
            self.category = '4'
        elif 2100 <= max(self.width, self.height):
            self.category = '5'
        else:
            print('Category could not be set')

    def strip_unit(self, areaValue):
        if areaValue.count('m') == 0:
            return areaValue
        else:
            return areaValue[0:-3]
