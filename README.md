# Export and Import Excel Template
  
  [![Ruby](https://badgen.net/badge/ruby/v2.7.0/:color?icon=ruby&color=red)](https://www.ruby-lang.org/en/news/2019/12/25/ruby-2-7-0-released) 
  [![Rails](https://badgen.net/badge/rails/v6.0.4.1/:color?color=red)](https://rubygems.org/gems/rails/versions/6.0.4.1)
  [![Gem Version](https://badgen.net/badge/roo/2.8.3/:color?color=yellow)](https://rubygems.org/gems/roo/versions/2.8.3)
  [![Gem Version](https://badgen.net/badge/rubyzip/1.3.0/:color?color=yellow)](https://rubygems.org/gems/rubyzip/versions/1.3.0)
  [![Gem Version](https://badgen.net/badge/caxlsx/3.1.1/:color?color=yellow)](https://rubygems.org/gems/caxlsx/versions/3.1.1)
  [![Gem Version](https://badgen.net/badge/caxlsx_rails/0.6.2/:color?color=yellow)](https://rubygems.org/gems/caxlsx_rails/versions/0.6.2)
  [![Bootstrap](https://badgen.net/badge/bootstrap/^5.1.3/:color?color=blue)](https://getbootstrap.com/docs/5.1/getting-started/introduction/)

## Prerequisites

This project setup steps expect following tools installed on the system.

- Ruby [2.7.0](https://www.ruby-lang.org/en/news/2019/12/25/ruby-2-7-0-released)
- Rails [6.0.4.1](https://rubygems.org/gems/rails/versions/6.0.4.1)

## Installation

### 1. Clone the project

> #### Clone with SSH
```ruby
  git clone git@gitlab.com:hannisoe/excel-export-import.git
```
or
> #### Clone with HTTPS
```ruby
  git clone git@gitlab.com:hannisoe/excel-export-import.git
```

### 2. Go to the project directory

```ruby
  cd excel-export-import
```

### 3. Check your ruby version

```ruby
  ruby -v
``` 
The ouput should start with something like ruby 2.7.0

If not, install the right ruby version using ruby version management tools.

> #### Install using rvm
  Run the following command to install ruby with rvm.  
  ```shell
    rvm install ruby-2.7.0
  ```
> #### Install using rbenv
  Run the following command to install ruby with rbenv.
  ```shell
    rbenv install 2.7.0
   ```
   
### 4. Check your rails version

```ruby
  rails -v
``` 

The ouput should start with something like Rails 6.0.4.1.

If not, then run command: 
```ruby
  gem install rails -v 6.0.4.1
``` 

### 5. Install dependencies

Using [Bundler](https://github.com/bundler/bundler) and [Yarn](https://github.com/yarnpkg/yarn), run:
```ruby
  bundle install
  yarn install
```
or run:
```ruby
  bundle && yarn
```

## Initialize the database

### 1. Create and migrate the database

Run the following commands to create and migrate the database:
```ruby
  rails db:create 
  rails db:migrate 
```
or in order to do db:drop, db:create, db:migrate in single command line, run: 
```ruby
  rails db:migrate:reset
```

For active storage, run:
```ruby
  rails active_storage:install
  rails db:migrate
```

## Start the rails server

Start the rails server using the command given below:
```ruby
  rails server 
```
or
```ruby
  rails s
```

# Documentation

### 1. Create New Project

```ruby
  rails new project_name -d postgresql 
```
In project folder, run:
```ruby
  bundle install
  yarn install
```
For database configuration, in database.yml, set configuration with your username and password.
```ruby
  username: postgres
  password: root
  host: localhost
```
and run:
```ruby
  rails db:create
```

### 2. Scaffold an Item resource

```ruby
  rails g scaffold item 
```
### 3. Add columns in create_items migration file

```ruby
  def change
    create_table :items do |t|
      	t.string :brand
      	t.integer :price
    	t.integer :quantity
      	t.string :country
      	t.string :colour
    	t.string :company_name
      	t.string :model
      	t.text :other_spec
    	t.boolean :instock
  	t.timestamps
    end
  end 
```

### 4. Scaffold a Customer resource

```ruby
  rails g scaffold customer
```

### 5. Add columns in create_customers migration file

```ruby
  def change
    create_table :customers do |t|
      	t.references :item, null: false, foreign_key: true
      	t.string :ph_num
      	t.string :email
      	t.string :last_name
      	t.string :first_name
      	t.date :birthday
  	t.timestamps
    end
  end 
```
and run:
```ruby
  rails db:migrate
```

### 6. For active storage, run:

```ruby
  rails active_storage:install
  rails db:migrate
```

### 7. Add in customer.rb

```ruby
  belongs_to :item, foreign_key: "item_id"
  has_one_attached :image
```

### 8. For validation and table join scope, add in item.rb

```ruby
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
```

### 9. In config/routes.rb

```ruby
  resources :customers
  resources :items
  root to: 'items#index'
```

## Export Excel Template

### 10. Add in Gemfile

```ruby
  gem 'rubyzip'
  gem 'caxlsx'
  gem 'caxlsx_rails'
```
and run:
```ruby
  bundle install
```

### 11. In items_controller, add

```ruby
  def index
    @items = Item.join_customer_table.order('customers.item_id')
    respond_to do |format|
      	format.xlsx {
        response.headers[
        'Content-Disposition'
        ] = "attachment; filename=template.xlsx"
      	}
      	format.html { render :index }
    end
  end
```

### 12. Add a link to download the spreadsheet in index page (items/index.html.erb)

```erb
  <%= link_to 'Download Excel Template', items_path(format: :xlsx), :class => "btn btn btn-info text-white" %>
```

### 13. In items/index.html.erb

```erb
  <div class="container-fluid">
    <p id="notice"><%= notice %></p>

    <h1>Items</h1>

    <table class="table table-bordered">
        <thead>
          <tr>
            <th>Brand</th>
            <th>Country</th>
            <th>Price</th>
            <th>Quantity</th>
            <th>Customer Name</th>
            <th>Email</th>
            <th>Phone number</th>
            <th>Photo</th>
          </tr>
        </thead>

        <tbody>
        <% @items.each do |item| %>
            <tr>
            <td><%= item.brand %></td>
            <td><%= item.country %></td>
            <td><%= item.price %></td>
            <td><%= item.quantity %></td>
            <td><%= "#{item.customer.first_name}  #{item.customer.last_name}" %></td>
            <td><%= item.customer.email %></td>
            <td><%= item.customer.ph_num %></td>
            <td><%= image_tag item.customer.image , style: 'height:100px; width:auto;' %></td>
            </tr>
        <% end %>
        </tbody>
    </table>
    <br>
    <%= link_to 'Download Excel Template', items_path(format: :xlsx),:class => "btn btn btn-info text-white" %>
  </div>
```

### 14. Create template for new spreadsheet

run:
```ruby
  touch app/views/items/index.xlsx.axlsx
```

### 15. Add in index.xlsx.axlsx

```axlsx
  wb = xlsx_package.workbook

  wb.styles do |style|
    template_heading = style.add_style(b: true, sz: 16, alignment: {:horizontal => :center, :vertical => :center})
    data_heading = style.add_style(b: true, sz: 11,:border => { :style => :thin, :color => "00" }, bg_color: "FFFFCC", alignment: {:vertical => :center, :wrap_text => true})
    data = style.add_style(sz: 11,:border => {:style => :thin, :color => "00" }, alignment: {:vertical => :center, :wrap_text => true})
    
    right_thin_border =  style.add_style({:border => { :style => :thin, :color => '00', :name => :right, :edges => [:right] }})

    wb.add_worksheet(name: "item1") do |sheet|
        # Add a title row
        sheet.add_row ["Items"], style: template_heading

        # Add a blank row
        sheet.add_row []

        # Merge cells
        sheet.merge_cells("A1:I1")
        sheet.merge_cells("B3:C3")
        sheet.merge_cells("E3:F3")
        sheet.merge_cells("H3:I3")
        sheet.merge_cells("B4:C4")
        sheet.merge_cells("E4:F4")
        sheet.merge_cells("B5:C5")
        sheet.merge_cells("B6:C6")
        sheet.merge_cells("A7:A8")
        sheet.merge_cells("B7:I8")
        sheet.merge_cells("B9:D9")
        sheet.merge_cells("B10:D10")
        sheet.merge_cells("B11:D11")
        sheet.merge_cells("E10:E11")
        sheet.merge_cells("F9:F18")
        sheet.merge_cells("G9:I18")

        # Create row
        sheet.add_row ["Brand Name", nil, nil, "Price", nil, nil, "Quantity", nil, nil], style: [data_heading, data, data, data_heading, data, data, data_heading, data, data]
        sheet.add_row ["Model", nil, nil, "Colour", nil, nil, nil, nil, nil], style: [data_heading, data, data, data_heading, data, data, nil, nil, right_thin_border]
        sheet.add_row ["Company Name", nil, nil, nil, nil, nil, nil, nil, nil], style: [data_heading, data, data, nil, nil, nil, nil, nil, right_thin_border]
        sheet.add_row ["Made in", nil, nil, nil, nil, nil, nil, nil, nil], style: [data_heading, data, data, nil, nil, nil, nil, nil, right_thin_border]
        sheet.add_row ["Other Specifications", nil, nil, nil, nil, nil, nil, nil, nil], style: [data_heading, data, data, data, data, data, data, data, data]
        sheet.add_row [nil, nil, nil, nil, nil, nil, nil, nil, nil], style: [data, nil, nil, nil, nil, nil, nil, nil, right_thin_border]
        sheet.add_row ["Customer Name", nil, nil, nil, "Age", "Image", nil, nil, nil], style: [data_heading, data, data, data, data_heading,data_heading, data, data, data]
        sheet.add_row ["Email", nil, nil, nil, nil,nil, nil, nil,nil], style: [data_heading, data, data, data, data, data, nil, nil, right_thin_border]
        sheet.add_row ["Phone Number", nil, nil, nil, nil, nil, nil, nil,nil], style: [data_heading, data, data, data, data, data, nil, nil, right_thin_border]
        sheet.add_row [nil, nil, nil, nil, nil, nil, nil, nil, nil], style: [nil, nil, nil, nil, nil, data, nil, nil, right_thin_border]
        sheet.add_row [nil, nil, nil, nil, nil, nil, nil, nil, nil], style: [nil, nil, nil, nil, nil, data, nil, nil, right_thin_border]
        sheet.add_row [nil, nil, nil, nil, nil, nil, nil, nil, nil], style: [nil, nil, nil, nil, nil, data, nil, nil, right_thin_border]
        sheet.add_row [nil, nil, nil, nil, nil, nil, nil, nil, nil], style: [nil, nil, nil, nil, nil, data, nil, nil, right_thin_border]
        sheet.add_row [nil, nil, nil, nil, nil, nil, nil, nil, nil], style: [nil, nil, nil, nil, nil, data, nil, nil, right_thin_border]
        sheet.add_row [nil, nil, nil, nil, nil, nil, nil, nil, nil], style: [nil, nil, nil, nil, nil, data, nil, nil, right_thin_border]
        sheet.add_row [nil, nil, nil, nil, nil, nil, nil, nil, nil], style: [nil, nil, nil, nil, nil, data, data, data, data]

        sheet.column_widths 24, 12, 12, 12, 12, 12, 12 , 10, 10
    end
  end
```

## Import Excel

### 16. Add in Gemfile

```ruby
  gem "roo", "~> 2.8.0"
```
and run:
```ruby
  bundle install
```

### 17. In config/routes.rb

```ruby
  resources :items do
    collection do
        post :import
    end
  end
```

### 18. Add in item.rb

```ruby
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
```

### 19. Add in items/index.html.erb

```erb
  <h4>Import Items</h4> 
  <%= form_with url: import_items_path, local: true do |form| %>
    <%= form.file_field :file %>
    <%= form.submit "Import" %>
  <% end %>
```

### 20. Add import method in items_controller

```ruby
  def import
    Item.item_import(params[:file])
    redirect_to items_url
  end
```
