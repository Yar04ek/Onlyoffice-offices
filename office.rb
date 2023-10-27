require 'selenium-webdriver'
require 'csv'


options = Selenium::WebDriver::Firefox::Options.new
options.add_argument('--headless') # Запустить браузер в режиме headless


driver = Selenium::WebDriver.for :firefox, options: options

begin

  driver.get('https://www.onlyoffice.com')


  about_nav_item = driver.find_element(id: 'navitem_about')
  2.times { about_nav_item.click }


  about_menu = driver.find_element(id: 'navitem_about_menu')
  wait = Selenium::WebDriver::Wait.new(timeout: 10)
  wait.until { about_menu.displayed? }


  driver.execute_script('document.querySelector("#navitem_about_contacts").click()')


  sleep(2)


  new_tab_handle = driver.window_handles.last
  driver.switch_to.window(new_tab_handle)


  office_elements = driver.find_elements(css: '.companydata')

  offices = []

  office_elements.each do |element|
    office_data = {}

    begin
      region = element.find_element(css: '.region')
      office_data['Region'] = region.text
    rescue Selenium::WebDriver::Error::NoSuchElementError
    end

    begin
      company_name = element.find_element(css: 'b')
      office_data['CompanyName'] = company_name.text
    rescue Selenium::WebDriver::Error::NoSuchElementError
    end

    begin
      street_address = element.find_element(css: '[itemprop="streetAddress"]')
      office_data['StreetAddress'] = street_address.text
    rescue Selenium::WebDriver::Error::NoSuchElementError
    end

    begin
      address_country = element.find_element(css: '[itemprop="addressCountry"]')
      office_data['AddressCountry'] = address_country.text
    rescue Selenium::WebDriver::Error::NoSuchElementError
    end

    begin
      postal_code = element.find_element(css: '[itemprop="postalCode"]')
      office_data['PostalCode'] = postal_code.text
    rescue Selenium::WebDriver::Error::NoSuchElementError
    end

    begin
      telephone = element.find_element(css: '[itemprop="telephone"]')
      office_data['Telephone'] = telephone.text
    rescue Selenium::WebDriver::Error::NoSuchElementError
    end

    offices.push(office_data)
  end


  driver.quit

  if offices.length > 0
    
    CSV.open('offices.csv', 'w') do |csv|
      csv << offices.first.keys # Записать заголовки
      offices.each do |office|
        csv << office.values
      end
    end

    puts 'CSV файл успешно записан.'
  else
    puts 'Данные об офисах не найдены.'
  end
rescue StandardError => e
  puts "Произошла ошибка: #{e.message}"
end
