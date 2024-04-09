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
        self._configure_logging()
        self._initialize_browser()
        self._results = []
        self._temp_folder = 'temp_folder'

    def _configure_logging(self):
        self._logger = logging.getLogger(__name__)
        self._logger.setLevel(logging.INFO)
        self._handler = logging.StreamHandler()
        self._handler.setLevel(logging.INFO)
        self._formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s')
        self._handler.setFormatter(self._formatter)
        self._logger.addHandler(self._handler)

    def _initialize_browser(self):
        self._driver = webdriver.Chrome()

    def _navigate_to_site(self, url):
        try:
            self._driver.get(url)
            self._logger.info(f"Navigated to {url}")
        except Exception as e:
            self._logger.error(
                f"An error occurred while navigating to {url}: {
                    str(e)}"
            )

    def _search_news(self, search_phrase, category):
        self._logger.info(
            f"Searching for: {search_phrase} in category: {category}")
        try:
            search_input = self._driver.find_element(
                By.CSS_SELECTOR, "input[name='p']")
            search_input.send_keys(search_phrase)
            search_input.submit()

            WebDriverWait(
                self._driver, 10).until(
                EC.number_of_windows_to_be(2))
            self._driver.switch_to.window(self._driver.window_handles[-1])
            self._logger.info("Switched to new tab")

            self._select_category(category)
        except Exception as e:
            self._logger.error(f"An error occurred while searching: {str(e)}")

    def _download_image(self, url, filename):
        try:
            response = requests.get(url)
            if response.status_code == 200:
                with open(filename, 'wb') as f:
                    f.write(response.content)
                return filename
            else:
                self._logger.error(
                    f"Error downloading image: {
                        response.status_code}"
                )
                return None
        except Exception as e:
            self._logger.error(f"Error downloading image: {str(e)}")
            return None

    def _store_results_to_excel(self, search_phrase, category):
        try:
            df = pd.DataFrame(self._results)

            filename = f"{search_phrase}_{category}_results.xlsx"
            df.to_excel(filename, index=False)
            self._logger.info(f"Results saved to {filename}")
        except Exception as e:
            self._logger.error(f"Error storing results: {str(e)}")

    def _select_category(self, category):
        try:
            category_switch = {
                'news': self._select_latest_news,
                'images': self._select_latest_image,
                'videos': self._select_latest_video,
                'local': self._select_latest_local,
                'shopping': self._select_latest_shopping
            }

            category_links = self._driver.find_elements(
                By.CSS_SELECTOR, "div.dd.horizontal-pivots a.td-n")
            available_categories = [link.text.strip().lower()
                                    for link in category_links]

            if category.lower() in available_categories:
                category_link = category_links[available_categories.index(
                    category.lower())]
                category_link.click()
                self._logger.info(f"Selected category: {category}")

                selected_category_method = category_switch.get(
                    category.lower())
                if selected_category_method:
                    latest_category_elements = selected_category_method()
                    if latest_category_elements:
                        for element in latest_category_elements:
                            self._extract_category_details(
                                category.lower(), element)
                return
            elif "more" in available_categories:
                more_link = category_links[available_categories.index("more")]
                more_link.click()
                self._logger.info(
                    "Clicked on 'More' to reveal hidden categories")
                return self._select_category(category)
            else:
                self._logger.warning(
                    f"Category '{category}' not found. Available categories: {
                                     available_categories}"
                )
        except Exception as e:
            self._logger.error(
                f"Error selecting category {category}: {
                    str(e)}"
            )

    def _extract_category_details(self, category, latest_category_element):
        if category == 'news':
            self._extract_news_details(latest_category_element)
        elif category == 'videos':
            self._extract_video_details(latest_category_element)
        elif category == 'images':
            self._extract_image_details(latest_category_element)
        elif category == 'local':
            self._extract_local_details(latest_category_element)
        elif category == 'shopping':
            self._extract_shopping_details(latest_category_element)

    def _select_latest_news(self):
        try:
            latest_news_elements = self._driver.find_elements(
                By.CSS_SELECTOR, "div.dd.NewsArticle")
            if latest_news_elements:
                self._logger.info("Selected the latest news")
                return latest_news_elements
            else:
                self._logger.info("No latest news found")
                return None
        except Exception as e:
            self._logger.error(f"Error selecting latest news: {str(e)}")
            return None

    def _extract_news_details(self, latest_news_element):
        try:
            title_element = latest_news_element.find_element(
                By.CSS_SELECTOR, "h4.s-title a")
            news_title = title_element.text.strip()
            self._logger.info(f"Title: {news_title}")

            source_element = latest_news_element.find_element(
                By.CSS_SELECTOR, ".s-source")
            news_source = source_element.text.strip()
            self._logger.info(f"Source: {news_source}")

            time_element = latest_news_element.find_element(
                By.CSS_SELECTOR, ".s-time")
            news_time = time_element.text.strip()
            self._logger.info(f"Time: {news_time}")

            description_element = latest_news_element.find_element(
                By.CSS_SELECTOR, ".s-desc")
            news_description = description_element.text.strip()
            self._logger.info(f"Description: {news_description}")

            url = title_element.get_attribute("href")
            self._logger.info(f"URL: {url}")

            self._results.append({'Title': news_title,
                                  'Source': news_source,
                                  'Time': news_time,
                                  'Description': news_description,
                                  'URL': url})
        except Exception as e:
            self._logger.error(f"Error extracting news details: {str(e)}")

    def _select_latest_image(self):
        try:
            image_elements = self._driver.find_elements(
                By.CSS_SELECTOR, "li.ld.r0")
            if image_elements:
                self._logger.info("Found images in the latest images section")
                return image_elements
            else:
                self._logger.info(
                    "No images found in the latest images section")
                return None
        except Exception as e:
            self._logger.error(f"Error selecting latest images: {str(e)}")
            return None

    def _extract_image_details(self, image_element):
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
            image_filename = os.path.join(
                self._temp_folder, f"{
                    re.sub(
                        '[^\\w\\-_. ]', '_', image_title)}.jpg"
            )
            if not os.path.exists(self._temp_folder):
                os.makedirs(self._temp_folder)
            self._download_image(image_url, image_filename)

            self._results.append({
                'Title': image_title,
                'URL': url,
                'Source': image_source,
                'Image Filename': image_filename
            })
        except Exception as e:
            self._logger.error(f"Error extracting image details: {str(e)}")

    def _select_latest_video(self):
        try:
            video_elements = self._driver.find_elements(
                By.CSS_SELECTOR, "li.vr.vres")
            if video_elements:
                self._logger.info("Found videos in the latest videos section")
                return video_elements
            else:
                self._logger.info(
                    "No videos found in the latest videos section")
                return None
        except Exception as e:
            self._logger.error(f"Error selecting latest video: {str(e)}")
            return None

    def _extract_video_details(self, video_element):
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
            self._logger.error(f"Error extracting video details: {str(e)}")
            return None

    def _select_latest_local(self):
        try:
            local_elements = self._driver.find_elements(
                By.CSS_SELECTOR, "li.list")
            if local_elements:
                self._logger.info("Found local results")
                return local_elements
            else:
                self._logger.info("No local results found")
                return None
        except Exception as e:
            self._logger.error(
                f"Error selecting latest local result: {
                    str(e)}"
            )
            return None

    def _extract_local_details(self, local_element):
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
            self._logger.error(f"Error extracting local details: {str(e)}")
            return None

    def _select_latest_shopping(self):
        try:
            shopping_elements = self._driver.find_elements(
                By.CSS_SELECTOR, "li.Grid__Cell-sc-1xfj9j2-1.exzZYi")
            if shopping_elements:
                self._logger.info("Found shopping results")
                return shopping_elements
            else:
                self._logger.info("No shopping results found")
                return None
        except Exception as e:
            self._logger.error(
                f"Error selecting latest shopping result: {
                    str(e)}"
            )
            return None

    def _extract_shopping_details(self, shopping_element):
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
            self._logger.error(f"Error extracting shopping details: {str(e)}")
            return None

    def _close_driver(self):
        self._driver.quit()
        logging.shutdown()

    def run(self, search_phrase, category):
        try:
            self._navigate_to_site("https://news.yahoo.com/")
            self._search_news(search_phrase, category)
            self._store_results_to_excel(search_phrase, category)
        except Exception as e:
            self._logger.error(f"An error occurred: {str(e)}")
        finally:
            self._close_driver()


if __name__ == "__main__":
    search_phrase = input("Enter search phrase: ")
    category = input("Enter category: ")

    news_bot = NewsBot()
    news_bot.run(search_phrase, category)
