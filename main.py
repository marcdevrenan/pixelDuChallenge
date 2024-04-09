import os
import pandas as pd
import re
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging


class NewsBot:
    def __init__(self):
        self.configure_logging()
        self.initialize_browser()
        self.results = []
        self.temp_folder = 'temp_folder'

    def configure_logging(self):
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)
        self.handler = logging.StreamHandler()
        self.handler.setLevel(logging.INFO)
        self.formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s')
        self.handler.setFormatter(self.formatter)
        self.logger.addHandler(self.handler)

    def initialize_browser(self):
        self.driver = webdriver.Chrome()

    def navigate_to_site(self):
        try:
            self.driver.get("https://news.yahoo.com/")
            self.logger.info("Navigated to Yahoo! News website")
        except Exception as e:
            self.logger.error(
                f"An error occurred while navigating to the website: {
                    str(e)}")

    def search_news(self, search_phrase, category):
        self.logger.info(
            f"Searching for: {search_phrase} in category: {category}")
        try:
            search_input = self.driver.find_element(
                By.CSS_SELECTOR, "input[name='p']")
            search_input.send_keys(search_phrase)
            search_input.submit()

            WebDriverWait(self.driver, 10).until(EC.number_of_windows_to_be(2))
            self.driver.switch_to.window(self.driver.window_handles[-1])
            self.logger.info("Switched to new tab")

            self.select_category(category)
        except Exception as e:
            self.logger.error(f"An error occurred while searching: {str(e)}")

    def download_image(self, url, filename):
        try:
            response = requests.get(url)
            if response.status_code == 200:
                with open(filename, 'wb') as f:
                    f.write(response.content)
                return filename
            else:
                self.logger.error(
                    f"Error downloading image: {
                        response.status_code}")
                return None
        except Exception as e:
            self.logger.error(f"Error downloading image: {str(e)}")
            return None

    def store_results_to_excel(self, search_phrase, category):
        try:
            df = pd.DataFrame(self.results)

            filename = f"{search_phrase}_{category}_results.xlsx"
            df.to_excel(filename, index=False)
            self.logger.info(f"Results saved to {filename}")
        except Exception as e:
            self.logger.error(f"Error storing results: {str(e)}")

    def select_category(self, category):
        try:
            category_switch = {
                'news': self.select_latest_news,
                'images': self.select_latest_image,
                'videos': self.select_latest_video,
                'local': self.select_latest_local,
                'shopping': self.select_latest_shopping
            }

            category_links = self.driver.find_elements(
                By.CSS_SELECTOR, "div.dd.horizontal-pivots a.td-n")
            available_categories = [link.text.strip().lower()
                                    for link in category_links]

            if category.lower() in available_categories:
                category_link = category_links[available_categories.index(
                    category.lower())]
                category_link.click()
                self.logger.info(f"Selected category: {category}")

                selected_category_method = category_switch.get(
                    category.lower())
                if selected_category_method:
                    latest_category_elements = selected_category_method()
                    if latest_category_elements:
                        for element in latest_category_elements:
                            self.extract_category_details(
                                category.lower(), element)
                return
            elif "more" in available_categories:
                more_link = category_links[available_categories.index("more")]
                more_link.click()
                self.logger.info(
                    "Clicked on 'More' to reveal hidden categories")
                return self.select_category(category)
            else:
                self.logger.warning(
                    f"Category '{category}' not found. Available categories: {
                        available_categories}"
                )
        except Exception as e:
            self.logger.error(f"Error selecting category {category}: {str(e)}")

    def extract_category_details(self, category, latest_category_element):
        if category == 'news':
            self.extract_news_details(latest_category_element)
        elif category == 'videos':
            self.extract_video_details(latest_category_element)
        elif category == 'images':
            self.extract_image_details(latest_category_element)
        elif category == 'local':
            self.extract_local_details(latest_category_element)
        elif category == 'shopping':
            self.extract_shopping_details(latest_category_element)

    def select_latest_news(self):
        try:
            latest_news_elements = self.driver.find_elements(
                By.CSS_SELECTOR, "div.dd.NewsArticle")
            if latest_news_elements:
                self.logger.info("Selected the latest news")
                return latest_news_elements
            else:
                self.logger.info("No latest news found")
                return None
        except Exception as e:
            self.logger.error(f"Error selecting latest news: {str(e)}")
            return None

    def extract_news_details(self, latest_news_element):
        try:
            title_element = latest_news_element.find_element(
                By.CSS_SELECTOR, "h4.s-title a")
            news_title = title_element.text.strip()
            self.logger.info(f"Title: {news_title}")

            source_element = latest_news_element.find_element(
                By.CSS_SELECTOR, ".s-source")
            news_source = source_element.text.strip()
            self.logger.info(f"Source: {news_source}")

            time_element = latest_news_element.find_element(
                By.CSS_SELECTOR, ".s-time")
            news_time = time_element.text.strip()
            self.logger.info(f"Time: {news_time}")

            description_element = latest_news_element.find_element(
                By.CSS_SELECTOR, ".s-desc")
            news_description = description_element.text.strip()
            self.logger.info(f"Description: {news_description}")

            url = title_element.get_attribute("href")
            self.logger.info(f"URL: {url}")

            self.results.append({'Title': news_title,
                                 'Source': news_source,
                                 'Time': news_time,
                                 'Description': news_description,
                                 'URL': url})
        except Exception as e:
            self.logger.error(f"Error extracting news details: {str(e)}")

    def select_latest_image(self):
        try:
            image_elements = self.driver.find_elements(
                By.CSS_SELECTOR, "li.ld.r0")
            if image_elements:
                self.logger.info("Found images in the latest images section")
                return image_elements
            else:
                self.logger.info(
                    "No images found in the latest images section")
                return None
        except Exception as e:
            self.logger.error(f"Error selecting latest images: {str(e)}")
            return None

    def extract_image_details(self, image_element):
        try:
            title_element = image_element.find_element(
                By.CSS_SELECTOR, "div.img-desc span.title")
            image_title = title_element.text.strip()

            url = image_element.find_element(
                By.CSS_SELECTOR, "a.redesign-img").get_attribute("href")

            source_element = image_element.find_element(
                By.CSS_SELECTOR, "div.img-desc span.source")
            image_source = source_element.text.strip()

            image_url = image_element.find_element(
                By.CSS_SELECTOR, "img").get_attribute("src")
            image_filename = os.path.join(self.temp_folder, f"{re.sub(
                '[^\\w\\-_. ]', '_', image_title)}.jpg")
            if not os.path.exists(self.temp_folder):
                os.makedirs(self.temp_folder)
            self.download_image(image_url, image_filename)

            self.results.append({
                'Title': image_title,
                'URL': url,
                'Source': image_source,
                'Image Filename': image_filename
            })
        except Exception as e:
            self.logger.error(f"Error extracting image details: {str(e)}")

    def select_latest_video(self):
        try:
            video_elements = self.driver.find_elements(
                By.CSS_SELECTOR, "li.vr.vres")
            if video_elements:
                self.logger.info("Found videos in the latest videos section")
                return video_elements
            else:
                self.logger.info(
                    "No videos found in the latest videos section")
                return None
        except Exception as e:
            self.logger.error(f"Error selecting latest video: {str(e)}")
            return None

    def extract_video_details(self, video_element):
        try:
            title_element = video_element.find_element(By.CSS_SELECTOR, "h3")
            video_title = title_element.text.strip()

            url = video_element.find_element(
                By.CSS_SELECTOR, "a").get_attribute("href")

            source_element = video_element.find_element(
                By.CSS_SELECTOR, "cite.url")
            video_source = source_element.text.strip()

            age_element = video_element.find_element(
                By.CSS_SELECTOR, "div.v-age")
            video_age = age_element.text.strip()

            return {
                'Title': video_title,
                'URL': url,
                'Source': video_source,
                'Age': video_age
            }
        except Exception as e:
            self.logger.error(f"Error extracting video details: {str(e)}")
            return None

    def select_latest_local(self):
        try:
            local_elements = self.driver.find_elements(
                By.CSS_SELECTOR, "li.list")
            if local_elements:
                self.logger.info("Found local results")
                return local_elements
            else:
                self.logger.info("No local results found")
                return None
        except Exception as e:
            self.logger.error(f"Error selecting latest local result: {str(e)}")
            return None

    def extract_local_details(self, local_element):
        try:
            title_element = local_element.find_element(
                By.CSS_SELECTOR, "div.titlewrapper a")
            local_title = title_element.text.strip()

            url = title_element.get_attribute("href")

            address_element = local_element.find_element(
                By.CSS_SELECTOR, "span.addr")
            local_address = address_element.text.strip()

            phone_element = local_element.find_element(
                By.CSS_SELECTOR, "span.hoo")
            local_phone = phone_element.text.strip()

            return {
                'Title': local_title,
                'URL': url,
                'Address': local_address,
                'Phone': local_phone
            }
        except Exception as e:
            self.logger.error(f"Error extracting local details: {str(e)}")
            return None

    def select_latest_shopping(self):
        try:
            shopping_elements = self.driver.find_elements(
                By.CSS_SELECTOR, "li.Grid__Cell-sc-1xfj9j2-1.exzZYi")
            if shopping_elements:
                self.logger.info("Found shopping results")
                return shopping_elements
            else:
                self.logger.info("No shopping results found")
                return None
        except Exception as e:
            self.logger.error(
                f"Error selecting latest shopping result: {
                    str(e)}"
            )
            return None

    def extract_shopping_details(self, shopping_element):
        try:
            title_element = shopping_element.find_element(
                By.CSS_SELECTOR,
                "span.FluidProductCell__Title-sc-fsx0f7-9.iKSKkx")
            shopping_title = title_element.text.strip()

            url = shopping_element.find_element(
                By.CSS_SELECTOR,
                "a.unstyled-link").get_attribute("href")

            price_element = shopping_element.find_element(
                By.CSS_SELECTOR,
                "span.FluidProductCell__PriceText-sc-fsx0f7-10.crhPQI")
            shopping_price = price_element.text.strip()

            merchant_element = shopping_element.find_element(
                By.CSS_SELECTOR,
                "span.FluidProductCell__MerchantInfo-sc-fsx0f7-8.dycvEi")
            shopping_merchant = merchant_element.text.strip()

            return {
                'Title': shopping_title,
                'URL': url,
                'Price': shopping_price,
                'Merchant': shopping_merchant
            }
        except Exception as e:
            self.logger.error(f"Error extracting shopping details: {str(e)}")
            return None

    def close_driver(self):
        self.driver.quit()


if __name__ == "__main__":
    search_phrase = input("Enter search phrase: ")
    category = input("Enter category: ")

    news_bot = NewsBot()
    try:
        news_bot.navigate_to_site()
        news_bot.search_news(search_phrase, category)
        news_bot.store_results_to_excel(search_phrase, category)
    except Exception as e:
        logging.error(f"An error occurred: {str(e)}")
    finally:
        news_bot.close_driver()
