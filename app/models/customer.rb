class Customer < ApplicationRecord
    belongs_to :item, foreign_key: "item_id"
    has_one_attached :image
end
