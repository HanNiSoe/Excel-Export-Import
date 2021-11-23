class Item < ApplicationRecord 
    validates_presence_of :brand
    validates_presence_of :price
    validates_presence_of :quantity
    validates_presence_of :country
    validates_presence_of :colour
    validates_presence_of :company_name
    validates_presence_of :model
    validates_presence_of :other_spec

    has_one :customer, dependent: :destroy   
    scope :join_customer_table, -> {joins("INNER JOIN customers ON customers.item_id = items.id")}

    def self.item_import(file)
        sheet = self.open_sheet(file)
        if sheet
            (sheet.sheets).map do |i|
                brand = ['brand', sheet.cell(3, 'B', i)]
                price = ['price', sheet.cell(3,'E',i)]
                quantity = ['quantity', sheet.cell(3, 'H', i)]
                country = ['country', sheet.cell(6, 'B', i)]
                colour = ['colour', sheet.cell(4, 'E', i)]
                company_name = ['company_name', sheet.cell(5, 'B', i)]
                model = ['model', sheet.cell(4, 'B', i)]
                other_spec = ['other_spec', sheet.cell(7, 'B', i)]
                item_data = Hash[[brand, price, quantity, country, colour, company_name, model, other_spec]]
                
                images = sheet.images(i)[0]
                name = sheet.cell(9, 'B', i).present? ? sheet.cell(9, 'B', i).split(' ') : []
                last_name = ['last_name' , name[1]]
                first_name = ['first_name' , name[0]]
                ph_num = ['ph_num' , sheet.cell(11, 'B', i)]
                email = ['email' , sheet.cell(10, 'B', i)]
                customer_data = Hash[[last_name, first_name, ph_num, email]]
    
                create_item(item_data, customer_data, images)
            end
        else
            false
        end
    end

    def self.create_item(data1, data2, image = nil)
        item = Item.find_by(brand: data1['brand']) || new
        item.attributes = data1
        if item.save
            if Customer.find_by(item_id: item.id)
                item.customer.update(data2)
            else
                item.create_customer(data2)
            end
            item.customer.image.attach(io: File.open(image),filename: "customer#{data2['last_name']}.jpg", content_type: 'image/jpeg') if image.present?
        else
            false
        end
    end

    def self.open_sheet(file)
        if file.present?
            case File.extname(file.original_filename)
            when ".csv" then Csv.new(file.path, nil, :ignore)
            when ".xls" then Roo::Excel.new(file.path, nil, :ignore)
            when ".xlsx" then Roo::Excelx.new(file.path)
            else raise "Unknown file type: #{file.original_filename}"
            end
        else
            false
        end
    end
end
