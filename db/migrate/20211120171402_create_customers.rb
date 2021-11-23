class CreateCustomers < ActiveRecord::Migration[6.0]
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
end
