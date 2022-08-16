
def get_valid_input(input_string,valid_options):
    input_string += "({})".format(", ".join(valid_options))
    response = input(input_string)
    while response.lower() not in valid_options:
        response = input(input_string)
    return response


class Property:
    def __init__(self,square_feet="",beds='',baths='',**kwargs):
        super().__init__(**kwargs)
        self.square_feet = square_feet
        self.num_bedrooms = beds
        self.num_baths = baths

    def display(self):
        print("PROPERTY DETAILS")
        print('================')
        print('square footage: {}'.format(self.square_feet))
        print('bedrooms: {}'.format(self.num_bedrooms))
        print('bathrooms: {}'.format(self.num_baths))

    # @staticmethod
    def prompt_init():
        return dict(square_feet = input("Enter the square feet"),
                    beds = input("Enter number of bedrooms:"),
                    baths = input("Enter number of baths"))

    prompt_init = staticmethod(prompt_init)


class Apartment(property):

    valid_laundries = ("coin","ensuite","none")
    valid_balconies = ("yes","no","solarium")

    def __init__(self,balcony='',laundry='',**kwargs):
        super().__init__(**kwargs)
        self.balcony = balcony
        self.laundry = laundry

    def display(self):
        super().display()
        print("APARTMENT DETAILS")
        print("Laundry: %s" % self.laundry)
        print("has balcony: %s" % self.balcony)

    def prompt_init():
        parent_init = Property.prompt_init()
        laundry = get_valid_input("What laundry facilities does the property have?",Apartment.valid_laundries)
        balcony = get_valid_input("Does the property have a balcony?",Apartment.valid_balconies)
        parent_init.update({
            "laundry":laundry,
            "balcony":balcony
        })
        return parent_init

    prompt_init = staticmethod(prompt_init)

class House(Property):
    valid_garage = ("attached","detached",'none')
    valid_fenced = ("yes","no")


    def __init__(self,num_stories="",garage="",fenced="",**kwargs):
        super().__init__(**kwargs)
        self.num_stories = num_stories
        self.garage = garage
        self.fenced = fenced

    def display(self):
        super().display()
        print("HOUSE DETIALS")
        print("number of stories: %s" % self.num_stories)
        print("garage :%s" % self.garage)
        print("fenced :%s" % self.fenced)

    def prompt_init():
        parent_init = Property.prompt_init()
        garage = get_valid_input("Does the property have a garage?",House.valid_garage)
        fenced = get_valid_input("Is the property fenced?", House.valid_fenced)
        num_stories = input("How many stories?")
        parent_init.update({
            "garage": garage,
            "fenced":fenced,
            "num_stories":num_stories
        })
        return parent_init

    prompt_init = staticmethod(prompt_init)

class Rental:

    def __init__(self,utilities='',furnished='',rent='',**kwargs):
        super().__init__(**kwargs)
        self.utilities = utilities
        self.furnished = furnished
        self.rent = rent

    def display(self):
        super().display()
        print('RENTAL DETAILS')
        print("estimated utilities : %s" % self.utilities)
        print("furnished: %s" % self.furnished)
        print("rent: %s" % self.rent)

    def prompt_init():
        return dict(
            rent=input("What is the monthly rent?"),
            utilities = input("What are the estimated utilities"),
            furnished = get_valid_input("Is the property furnished?",("yes","no"))
        )

    prompt_init = staticmethod(prompt_init)

class Purchase:

    def __init__(self,purchase_price="",property_tax = "",**kwargs):
        super().__init__(**kwargs)
        self.purchase_price = purchase_price
        self.property_tax = property_tax

    def display(self):
        super().display()
        print("PURCHASE_DETAILS")
        print("purchase price : %s" % self.purchase_price)
        print("property_tax : %s" % self.property_tax)

    def prompt_init():
        return dict({
            "price" : input("What is the selling price?"),
            "tax" : input("What are the estimated taxes?")
        })
    prompt_init = staticmethod(prompt_init)


class HouseRental(House,Rental):
    def prompt_init():
        init = House.prompt_init()
        init.update(Rental.prompt_init())
        return init
    prompt_init = staticmethod(prompt_init)

class ApartmentRental(Apartment,Rental):
    def prompt_init():
        init = Apartment.prompt_init()
        init.update(Rental.prompt_init())
        return init

    prompt_init = staticmethod(prompt_init)

class HousePurchase(House,Purchase):
    def prompt_init():
        init = House.prompt_init()
        init.update(Purchase.prompt_init())
        return init
    prompt_init = staticmethod(prompt_init)


class ApartmentPurchase(Apartment,Purchase):
    def prompt_init():
        init = Apartment.prompt_init()
        init.update(Purchase.prompt_init())
        return init
    prompt_init = staticmethod(prompt_init)

class Agent:
    def __init__(self):
        self.property_list = []

    def display_properties(self):
        for property in self.property_list:
            property.display()

    type_map = {
        ("house","rental") : HouseRental,
        ("house","purchase") : HousePurchase,
        ("apartment","rental") : ApartmentRental,
        ("apartment","purchase") : ApartmentPurchase
    }

    def add_property(self):
        property_type = get_valid_input("What type of property? ", ("house","apartment")).lower()
        payment_type = get_valid_input("What payment type?", ("purchase","rental")).lower()
        PropertyClass = self.type_map[(property_type,payment_type)]
        init_args = PropertyClass.prompt_init()
        self.property_list.append(PropertyClass(**init_args))












