import os
from openai import OpenAI
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    ElementClickInterceptedException
)
import time
from webdriver_manager.chrome import ChromeDriverManager

# Load environment variables from the .env file
load_dotenv()

# Get credentials and API key from environment variables
user_id = os.getenv('NAVER_USER_ID')
password = os.getenv('NAVER_PASSWORD')
OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')

# Set OpenAI API key
client = OpenAI(api_key=OPENAI_API_KEY)

def generate_reply(review_content):
    system_prompt = """
    당신은 쇼핑몰 고객 문의에 응답하는 친절한 상담원입니다. 한국어 문법과 띄어쓰기를 정확히 지키십시오.
    ### 응답 가이드라인:
    이모지를 쓰지 마십시오. 지원되지 않습니다. TEXT로만 구성하십시오.

    1. **긍정적인 리뷰:**
    - 감사 인사를 전하며, 고객의 긍정적인 경험을 기뻐하는 메시지를 전달합니다.
    - 제품이나 서비스의 특징을 언급하고, 고객의 피드백이 중요하다는 점을 강조합니다.
    - 예시: "소중한 의견 남겨주셔서 감사합니다. 저희 제품을 만족스럽게 사용해주셔서 기쁩니다. 앞으로도 최선을 다하겠습니다."

    2. **중립적/개선 요청 리뷰:**
    - 감사 인사 후 불편에 공감하며, 문제 해결을 위한 조치를 설명합니다.
    - 추가 지원이 필요할 경우 고객 지원팀과의 연결을 안내합니다.
    - 예시: "남겨주신 의견 감사합니다. 말씀하신 사항을 중요하게 생각하며, 더 나은 서비스를 위해 노력하겠습니다."

    3. **부정적인 리뷰:**
    - 진심으로 사과하고, 불편 사항에 대해 공감합니다.
    - 문제 해결 방법을 안내하며, 추가 문의가 있을 경우 언제든지 지원을 안내합니다.
    - 예시: "불편을 드려 죄송합니다. 문제를 신속히 해결할 수 있도록 최선을 다하겠습니다. 추가 문의가 있으시면 언제든지 말씀해 주세요."

    4. **반복적인 리뷰:**
    - 유사한 내용이라도 개인화된 문구를 사용해 기계적인 응답처럼 보이지 않도록 합니다.

    ### 기본 톤과 스타일:
    - 항상 긍정적이고 정중하며, 고객의 입장에서 생각하는 태도를 보여줍니다.
    - 간결하고 깔끔하게 응대하며, 고객이 이해하기 쉬운 언어를 사용합니다.
    - 조사한 자료에 근거한 정확한 답변을 제공합니다.
    """
    user_prompt = f"고객의 리뷰에 답변하세요. {review_content}"
    chat_completion = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
    )
    return chat_completion.choices[0].message.content.strip()


def login_and_navigate(driver):
    # Navigate to the login page
    driver.get(login_url)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//input[@type="text" and @placeholder="아이디 또는 이메일 주소" and @class="Login_ipt__cPqIR"]'))
    ).send_keys(user_id)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//input[@type="password" and @placeholder="비밀번호" and @class="Login_ipt__cPqIR"]'))
    ).send_keys(password)
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//span[@class="Button_txt__c0W-8" and text()="로그인"]'))
    ).click()
    WebDriverWait(driver, 20).until(EC.url_contains('sell.smartstore.naver.com'))
    time.sleep(5)

    # Store selection process
    try:
        store_select_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//a[@role="button" and @ui-sref="work.channel-select" and @data-action-location-id="selectStore"]'))
        )
        store_select_button.click()
        kazmi_store_option = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//label[.//span[contains(text(), "카즈미")]]'))
        )
        kazmi_store_option.click()
        time.sleep(5)
    except Exception as e:
        print(f"Error during initial store selection: {e}")
    time.sleep(5)


# Define URLs
login_url = "https://accounts.commerce.naver.com/login?url=https%3A%2F%2Fsell.smartstore.naver.com%2F%23%2Flogin-callback"
review_url = "https://sell.smartstore.naver.com/#/review/search"

# Setup the browser with specific options
options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--incognito')
# Uncomment the line below if headless mode is desired during testing (for debugging, keep it commented)
# options.add_argument('--headless')

# Initialize the WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.maximize_window()

try:
    # Retry mechanism
    retry_attempts = 3
    attempt = 0

    while attempt < retry_attempts:
        login_and_navigate(driver)

        try:
            # Navigate to the review page
            driver.get(review_url)

            # Click 3 months button
            WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@type="button" and contains(text(), "3개월")]'))
            ).click()

            # Dropdown to select "답글미등록"
            try:
                dropdown_toggle_xpath = '//div[@class="selectize-input items not-full ng-valid ng-pristine has-options"]'
                dropdown_toggle = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, dropdown_toggle_xpath))
                )
                dropdown_toggle.click()

                reply_unregistered_xpath = '//div[@data-value="false" and @data-selectable and contains(text(), "답글미등록")]'
                reply_unregistered_option = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, reply_unregistered_xpath))
                )
                reply_unregistered_option.click()
                time.sleep(1)

            except TimeoutException:
                print("Dropdown or option to select '답글미등록' not found.")

            # Click Search button after modification
            try:
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//span[@class="content" and text()="검색"]'))
                ).click()
                time.sleep(3)
            except TimeoutException:
                print("Search button click failed after modification.")

            # Initial scroll to position elements properly
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            driver.execute_script("window.scrollBy(0, -300);")
            time.sleep(1)

            # Track whether reviews were processed
            reviews_processed = False

            # Loop through reviews dynamically
            review_index = 1
            while True:
                try:
                    review_content_xpath = f'(//a[contains(@ng-click, "vm.func.openReviewDetailModal") and contains(@class, "text-info")])[{review_index}]'
                    try:
                        review_content_link = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, review_content_xpath))
                        )
                        review_content_link.click()
                    except ElementClickInterceptedException as e:
                        print(f"Element click intercepted for review {review_index}, closing overlay.")
                        modal_close_button_xpath = '//button[@type="button" and @class="close" and @data-dismiss="modal"]'
                        try:
                            close_button = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, modal_close_button_xpath))
                            )
                            close_button.click()
                            time.sleep(1)
                        except TimeoutException:
                            print("Could not close modal, moving to next review.")
                        review_index += 1
                        continue

                    except TimeoutException:
                        print("No more reviews to process. Retrying.")
                        break

                    reviews_processed = True

                    # Proceed with generating and submitting the reply
                    Naver_review_xpath = '//p[@class="text-sub mg-top txt-detail" and @ng-bind-html="vm.data.reviewContent"]'
                    review_element = WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, Naver_review_xpath))
                    )
                    review_content = review_element.text.strip()
                    if not review_content:
                        raise ValueError("Review content is empty")

                    print(f"Review content: {review_content}")

                    # Generate the reply
                    generated_reply = generate_reply(review_content)
                    print(f"Generated reply: {generated_reply}")

                    # Enter the reply and submit
                    try:
                        answer_xpath = '//textarea[contains(@class, "form-control")]'
                        reply_field = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, answer_xpath))
                        )
                        reply_field.clear()
                        reply_field.send_keys(generated_reply)

                        # Correctly find and click the submit button
                        Naver_submit_button_xpath = '//button[span[text()="답글 등록"]]'
                        submit_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, Naver_submit_button_xpath))
                        )
                        submit_button.click()

                        # Handle any confirmation pop-up
                        confirm_button_xpath = '//button[@type="button" and contains(@class, "btn-primary") and @ng-click="ok()"]'
                        try:
                            confirm_button = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, confirm_button_xpath))
                            )
                            confirm_button.click()
                        except TimeoutException:
                            print("Confirmation not needed.")

                    except (TimeoutException, NoSuchElementException) as submit_exception:
                        print(f"Error during reply process: {submit_exception}")
                        continue

                    # Always close the modal before moving to the next review
                    modal_close_button_xpath = '//button[@type="button" and @class="close" and @data-dismiss="modal"]'
                    try:
                        close_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, modal_close_button_xpath))
                        )
                        close_button.click()
                        time.sleep(1)
                    except TimeoutException:
                        print("Could not find a close button for the modal, attempting to proceed anyway.")

                    time.sleep(2)
                    review_index += 1

                except ElementClickInterceptedException:
                    continue

            if not reviews_processed:  # If no reviews were processed
                attempt += 1
                if attempt < retry_attempts:
                    print(f"Retrying from login page: Attempt {attempt}")
                    time.sleep(3)
                else:
                    print("Exceeded retry attempts with no reviews processed.")
            else:
                break  # Exit retry loop if some reviews were processed

        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            break

finally:
    if driver:
        driver.quit()