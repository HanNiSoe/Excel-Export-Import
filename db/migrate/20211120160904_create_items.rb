class CreateItems < ActiveRecord::Migration[6.0]
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
end
