import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook


class AmazonScraper:
    def __init__(self, url, product, price_range):
        self.url = url
        self.product = product
        self.price_range = int(price_range)  # Convert price_range to integer
        self.driver = None
        self.wait = None

        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, 10)

        self.products = []

        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = f"{self.product} Amazon Data"
        self.ws.append(["Name", "URL", "Price", "Rating"])  # Add headers to the Excel sheet

    def search_product(self):
        self.driver.get(self.url)
        input_field = self.wait.until(EC.presence_of_element_located((By.ID, 'twotabsearchtextbox')))
        input_field.send_keys(self.product)

        submit_button = self.wait.until(EC.element_to_be_clickable((By.ID, 'nav-search-submit-button')))
        submit_button.click()

    def extract_product_details(self):
        result_list = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.s-main-slot.s-result-list.s-search-results.sg-row')))
        product_elements = result_list.find_elements(By.CSS_SELECTOR,
                                                     '[data-asin][data-component-type="s-search-result"]')

        for product_element in product_elements:
            try:
                product_url = product_element.find_element(By.CSS_SELECTOR, 'h2 a').get_attribute('href')
                product_name = product_element.find_element(By.CSS_SELECTOR, 'h2 span').text
                product_price_element = product_element.find_element(By.CSS_SELECTOR, '.a-price-whole')
                product_price = product_price_element.text if product_price_element else '0'

                # Check multiple selectors for the rating element
                product_rating = "No rating"
                rating_selectors = ['.a-icon-alt', '.a-declarative .a-row .a-size-small .a-size-base']
                for selector in rating_selectors:
                    try:
                        product_rating_element = product_element.find_element(By.CSS_SELECTOR, selector)
                        if product_rating_element:
                            product_rating = product_rating_element.get_attribute('textContent')
                            break
                    except:
                        continue

                product_price_int = int(product_price.replace(',', ''))  # Convert price to integer

                if product_price_int <= self.price_range:
                    self.products.append({
                        "name": product_name,
                        "url": product_url,
                        "price": product_price,
                        "rating": product_rating
                    })
                    self.ws.append([product_name, product_url, product_price, product_rating])
            except Exception as e:
                # Skip any product that does not have the necessary information
                continue

    def extract_data_from_url(self):
        while True:
            self.extract_product_details()
            try:
                next_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.s-pagination-next')))
                self.driver.execute_script("arguments[0].scrollIntoView();", next_button)
                next_button.click()
                self.wait.until(EC.staleness_of(next_button))
                time.sleep(0)  # Add a delay to prevent being blocked by Amazon
            except Exception as e:
                break

    def save_excel(self):
        self.wb.save(f"{self.product}.xlsx")

    def run_scraper(self):
        try:
            self.search_product()
            self.extract_data_from_url()
            self.save_excel()
        finally:
            self.driver.quit()


# Example usage:
if __name__ == "__main__":
    url = "https://www.amazon.in/"
    product_name = input("Enter the Name of the Product: ")
    expected_price = input("Enter the price at which the product You want to buy: ")
    scraper = AmazonScraper(url, product_name, expected_price)
    scraper.run_scraper()
