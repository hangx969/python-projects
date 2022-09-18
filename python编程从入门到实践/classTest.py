# class test P142 9-1

class Restaurant():

    def __init__(self, restaurant_name, food_type):
      self.restaurant_name = restaurant_name
      self.food_type = food_type
      self.numServed = 0

    
    def describe_restaurant(self):
        print("name: " + self.restaurant_name)
        print("type: " + self.food_type)

    def open_restaurant(self):
        print("opening")

    def setNumServed(self, num):
        self.numServed = num
    
    def addNumServed(self, num):
        self.numServed += num

# define a instance
res1 = Restaurant("feimao", "spicy")
res1.describe_restaurant()
res1.open_restaurant()
print(res1.numServed)
res1.numServed = 1
print(res1.numServed)

res1.setNumServed(10)
print(res1.numServed)

res1.addNumServed(20)
print(res1.numServed)


# class test P147 9-4
