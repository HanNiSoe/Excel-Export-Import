Rails.application.routes.draw do
  resources :customers
  resources :items do
    collection do
      post :import
    end
  end
  root to: 'items#index'
  # For details on the DSL available within this file, see https://guides.rubyonrails.org/routing.html
end
