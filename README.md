# XlsExporter

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'xls_exporter'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install xls_exporter

## Usage

```ruby
class User
  attr_accessor :id, :first_name, :last_name, :email
end


require 'xls_exporter'

XlsExporter.export do
  default_style horizontal: :center, vertical_align: center, text_wrap: true
  
  font_size = 10,
  words_in_line = 5
  
  filename 'your-file-name'

  add_sheet 'your-sheet-name'
  export_models collection, :id, { full_name: -> { first_name + last_name } }, { user_email: :email }
end
```

Result

| ID | Full Name      | User Email                    |
|----|----------------|-------------------------------|
| 1  | Zhora Zhukov   | victory_marshal@army.su       |
| 2  | Petr The First | borody@doloi.ru               | 
| 3  | Vladimir P.    | babushka_would_be@dedushka.ru |


## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake test` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`.

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/xls_exporter. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [Contributor Covenant](http://contributor-covenant.org) code of conduct.

## Articles about

* [Ruby + Excel. Autofit cell content](https://medium.com/@kalashnikovisme/ruby-excel-autofit-cell-content-c1cd1e329706)
