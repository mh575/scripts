from playwright.sync_api import sync_playwright

ua = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36"

with sync_playwright() as p: # get user credentials to continue: UofA portal, prompt user in terminal, login, access SDS/CHEMINFO/CHEMINDEX
    browser = p.chromium.launch(headless = False, args = ['--disable-blink-features=AutomationControlled'])
    context = browser.new_context(
        user_agent = ua, 
        viewport = {'width':1280, 'height':1024},
        device_scale_factor = 1
        ) 
    page = browser.new_page()
    page.goto("https://www.sigmaaldrich.com/CA/en/search")
    # page.get_by_text("Search Type")
    page.locator('#onetrust-accept-btn-handler').click()
    page.get_by_text('Search Type')
    page.locator('#type').click()
    # this is where you search by name or cas -> download pdf -> parse with pymupdf lib -> csv results
    browser.close()