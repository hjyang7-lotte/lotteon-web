"""
W Concept 크롤러 - 헤드리스 모듈
웹 인터페이스 통합을 위한 GUI 없는 크롤러
"""

import sys
import os
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass
import time
import pandas as pd
from datetime import datetime
from urllib.parse import urljoin
from playwright.sync_api import sync_playwright

# 카테고리 URL 매핑
CATEGORY_URLS = {
    "베스트탭 (메인)": "https://display.wconcept.co.kr/rn/best?displayCategoryType=10101&gnbType=Y",
    "전체": "https://display.wconcept.co.kr/rn/best?displayCategoryType=ALL&displaySubCategoryType=ALL&gnbType=Y",
    "의류": "https://display.wconcept.co.kr/rn/best?displayCategoryType=10101&displaySubCategoryType=ALL&gnbType=Y",
    "가방": "https://display.wconcept.co.kr/rn/best?displayCategoryType=10102&displaySubCategoryType=ALL&gnbType=Y",
    "신발": "https://display.wconcept.co.kr/rn/best?displayCategoryType=10103&displaySubCategoryType=ALL&gnbType=Y",
    "ACC": "https://display.wconcept.co.kr/rn/best?displayCategoryType=10104&displaySubCategoryType=ALL&gnbType=Y",
    "뷰티": "https://display.wconcept.co.kr/rn/best?displayCategoryType=10107&displaySubCategoryType=ALL&gnbType=Y",
    "키즈": "https://display.wconcept.co.kr/rn/best?displayCategoryType=10109&displaySubCategoryType=ALL&gnbType=Y"
}


class WConceptCrawler:
    def __init__(self):
        self.log_callback = None
        
    def log(self, message):
        """로그 출력"""
        if self.log_callback:
            self.log_callback(message)
        else:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
    
    def crawl_products(self, category, count=10, headless=True):
        """상품 크롤링 실행"""
        url = CATEGORY_URLS.get(category)
        if not url:
            self.log(f"Error: Unknown category '{category}'")
            return []
        
        self.log(f"Starting W Concept crawling for '{category}' (Count: {count})")
        self.log(f"Headless mode: {headless}")
        
        results = []
        browser = None
        
        try:
            with sync_playwright() as p:
                self.log("Launching browser...")
                browser = p.chromium.launch(
                    headless=headless, 
                    timeout=60000,
                    args=["--no-sandbox", "--disable-dev-shm-usage"]
                )
                self.log("✅ Browser launched successfully")
                
                context = browser.new_context(
                    user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
                    viewport={"width": 1920, "height": 1080},
                    permissions=[]
                )
                context.set_default_timeout(60000)
                
                page = context.new_page()
                page.set_default_timeout(60000)
                
                # 알림 권한 자동 거부
                page.add_init_script("""
                    if (navigator.permissions) {
                        navigator.permissions.query({name: 'notifications'}).then(function(result) {});
                    }
                    const originalRequestPermission = Notification.requestPermission;
                    Notification.requestPermission = function() {
                        return Promise.resolve('denied');
                    };
                """)
                
                self.log("Navigating to best products page...")
                page.goto(url, timeout=90000, wait_until="domcontentloaded")
                
                try:
                    page.wait_for_load_state("networkidle", timeout=30000)
                except:
                    self.log("networkidle wait failed, continuing...")
                
                time.sleep(2)
                
                # 팝업 닫기
                self._close_popups(page)
                
                # 상품 버튼 찾기
                self.log("Finding product elements...")
                button_selectors = [
                    "button.sc-d9bca83f-7.area-click[type='button']",
                    "button.area-click[type='button']",
                    "button.sc-d9bca83f-7[type='button']"
                ]
                
                product_items = None
                for selector in button_selectors:
                    try:
                        test_buttons = page.locator(selector)
                        if test_buttons.count() > 0:
                            product_items = test_buttons
                            self.log(f"Found {product_items.count()} products with selector: {selector}")
                            break
                    except:
                        continue
                
                if not product_items or product_items.count() == 0:
                    self.log("Error: No products found")
                    return []
                
                # 유효한 상품만 필터링 (검증 로직 완화)
                valid_indices = []
                total_count = product_items.count()
                self.log(f"Processing {total_count} items...")
                
                # 모든 발견된 항목을 유효한 것으로 간주하고 추출 시도
                for idx in range(min(total_count, count + 20)):
                    valid_indices.append(idx)
                
                actual_count = min(len(valid_indices), count)
                self.log(f"Attempting to collect {actual_count} products")
                
                # 상품 정보 수집
                for i in range(actual_count):
                    try:
                        valid_idx = valid_indices[i]
                        item = product_items.nth(valid_idx)
                        
                        item.scroll_into_view_if_needed(timeout=5000)
                        time.sleep(0.3)
                        
                        # JavaScript로 상품 정보 추출 (로직 개선)
                        product_data = item.evaluate("""
                            (element) => {
                                const data = {
                                    brand: null,
                                    title: null,
                                    price: null,
                                    review_count: null,
                                    like_count: null
                                };
                                
                                // Helper to clean text
                                const clean = (text) => text ? text.trim() : null;
                                
                                // 브랜드 추출 시도
                                // 1. 명시적 클래스
                                const brandEl = element.querySelector('.text.title');
                                if (brandEl) data.brand = clean(brandEl.textContent);
                                
                                // 2. 'brand' 클래스 포함 요소
                                if (!data.brand) {
                                    const b = element.querySelector('[class*="brand"], [class*="Brand"]');
                                    if (b) data.brand = clean(b.textContent);
                                }
                                
                                // 상품명 추출 시도
                                // 1. 명시적 클래스
                                const titleEl = element.querySelector('.text.detail');
                                if (titleEl) data.title = clean(titleEl.textContent);
                                
                                // 2. 'product' 나 'info' 관련 클래스
                                if (!data.title) {
                                    const t = element.querySelector('[class*="product"], [class*="name"], [class*="ellips"]');
                                    if (t) data.title = clean(t.textContent);
                                }
                                
                                // 가격 추출 (할인가 -> 정가 순)
                                const finalPriceEl = element.querySelector('.text.final-price strong');
                                if (finalPriceEl) data.price = clean(finalPriceEl.textContent);
                                
                                if (!data.price) {
                                    const priceEl = element.querySelector('[class*="price"] strong, strong[class*="price"]');
                                    if (priceEl) data.price = clean(priceEl.textContent);
                                }
                                
                                if (!data.price) {
                                    // 텍스트에서 숫자+원/comma 패턴 찾기
                                    const text = element.textContent;
                                    const priceMatch = text.match(/([0-9,]+)\s*원?/);
                                    if (priceMatch) data.price = priceMatch[1];
                                }
                                
                                // 리뷰 수
                                const reviewSpan = element.querySelector('span.review');
                                if (reviewSpan) {
                                    const cntSpan = reviewSpan.querySelector('span.cnt, span[class*="cnt"]');
                                    if (cntSpan) {
                                        const reviewText = clean(cntSpan.textContent);
                                        const match = reviewText?.match(/\\d+/);
                                        if (match) data.review_count = match[0];
                                    }
                                }
                                
                                // 좋아요 수
                                const likeSpan = element.querySelector('span.like');
                                if (likeSpan) {
                                    const cntSpan = likeSpan.querySelector('span.cnt, span[class*="cnt"]');
                                    if (cntSpan) {
                                        data.like_count = clean(cntSpan.textContent);
                                    }
                                }
                                
                                return data;
                            }
                        """)
                        
                        brand = product_data.get("brand") or ""
                        title = product_data.get("title") or ""
                        price = product_data.get("price") or "가격 정보 없음"
                        review_count = product_data.get("review_count") or "0"
                        like_count = product_data.get("like_count") or "0"
                        
                        # 필수 정보 없어도 우선 수집하고 로그 남김 (빈 값 허용)
                        if not brand and not title:
                            self.log(f"[{i+1}] Warning: Empty brand/title inferred. HTML might have changed.")
                        
                        self.log(f"[{i+1}/{actual_count}] {brand} - {title[:30]}...")
                        
                        # 상세 페이지 URL 추출 - 상품 정보에서 ItemCD(상품 ID) 찾기
                        detail_url = ""
                        try:
                            detail_url = item.evaluate("""
                                (button) => {
                                    // 1. GA4 클릭 이벤트나 커스텀 속성에서 찾기
                                    const card = button.closest('.product-item') || button;
                                    if (card) {
                                        const html = card.outerHTML;
                                        const itemCdMatch = html.match(/item[Cc]d['"]?\\s*[:=]\\s*['"]?(\\d{9})/);
                                        if (itemCdMatch && itemCdMatch[1]) return 'https://www.wconcept.co.kr/Product/' + itemCdMatch[1];
                                    }
                                    
                                    // 2. 이미지 주소에서 추출 (매우 신뢰도 높음)
                                    let img = button.querySelector('img') || (card ? card.querySelector('img') : null);
                                    if (img && img.src) {
                                        const matches = img.src.match(/\\/(\\d{9})(_|\\.jpg)/);
                                        if (matches && matches[1]) {
                                            return 'https://www.wconcept.co.kr/Product/' + matches[1];
                                        }
                                    }
                                    
                                    // 3. 버튼의 onclick 속성 등에서 9자리 숫자 찾기
                                    const allAttr = button.outerHTML;
                                    const numMatch = allAttr.match(/\\d{9}/);
                                    if (numMatch) return 'https://www.wconcept.co.kr/Product/' + numMatch[0];
                                    
                                    return null;
                                }
                            """)
                            
                            if detail_url:
                                self.log(f"  → Detail URL: {detail_url}")
                            else:
                                self.log(f"  → No itemCd found for this product")
                        except Exception as e:
                            self.log(f"  → URL extraction error: {e}")
                        
                        # 판매자 정보 추출
                        seller_info = {"판매자명": "", "사업자등록번호": "", "통신판매업신고": "", 
                                     "대표자명": "", "주소": "", "연락처": "", "이메일": ""}
                        
                        if detail_url and detail_url != "URL 수집 실패":
                            self.log(f"  → Extracting seller info...")
                            seller_info = self._extract_seller_info(page, detail_url)
                            self.log(f"  → Seller: {seller_info.get('판매자명', 'N/A')}")
                            
                            # 목록 페이지로 돌아가기
                            try:
                                page.goto(url, timeout=30000, wait_until="domcontentloaded")
                                time.sleep(1)
                                self._close_popups(page)
                                
                                # 상품 목록 다시 찾기
                                for selector in button_selectors:
                                    try:
                                        test_buttons = page.locator(selector)
                                        if test_buttons.count() > 0:
                                            product_items = test_buttons
                                            break
                                    except:
                                        continue
                            except Exception as e:
                                self.log(f"  → Error returning to list: {e}")
                        else:
                            self.log(f"  → Skipping seller info (no URL)")
                        
                        results.append({
                            "순위": i + 1,
                            "브랜드": brand,
                            "상품명": title,
                            "가격": price,
                            "리뷰수": review_count,
                            "좋아요수": like_count,
                            "상세페이지URL": detail_url or "URL 수집 실패",
                            "판매자명": seller_info.get("판매자명", ""),
                            "사업자등록번호": seller_info.get("사업자등록번호", ""),
                            "통신판매업신고": seller_info.get("통신판매업신고", ""),
                            "대표자명": seller_info.get("대표자명", ""),
                            "주소": seller_info.get("주소", ""),
                            "연락처": seller_info.get("연락처", ""),
                            "이메일": seller_info.get("이메일", "")
                        })
                        
                    except Exception as e:
                        self.log(f"[{i+1}] Error collecting product: {e}")
                        continue

                
                self.log(f"✅ Collected {len(results)} products")
                
        except Exception as e:
            self.log(f"Crawling error: {e}")
            import traceback
            self.log(f"Traceback: {traceback.format_exc()}")
        finally:
            if browser:
                try:
                    browser.close()
                except:
                    pass
        
        return results
    
    def save_to_excel(self, products, category, output_dir=None):
        """엑셀 파일로 저장"""
        if not products:
            self.log("No products to save")
            return None
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"wconcept_best_{category}_{timestamp}.xlsx"
        
        # 출력 디렉토리 설정
        if output_dir is None:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            output_dir = os.path.join(base_dir, "results")
        
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except:
                pass
        
        filepath = os.path.join(output_dir, filename)
        
        try:
            df = pd.DataFrame(products)
            df.to_excel(filepath, index=False, engine='openpyxl')
            self.log(f"✅ Saved to: {filepath}")
            return filepath
        except Exception as e:
            self.log(f"Error saving Excel: {e}")
            return None
    
    def _extract_seller_info(self, page, detail_url):
        """상세 페이지에서 판매자 정보 추출"""
        seller_info = {
            "판매자명": "",
            "사업자등록번호": "",
            "통신판매업신고": "",
            "대표자명": "",
            "주소": "",
            "연락처": "",
            "이메일": ""
        }
        
        try:
            # 상세 페이지로 이동
            page.goto(detail_url, timeout=30000, wait_until="domcontentloaded")
            time.sleep(1.5)
            
            # 팝업 닫기
            self._close_popups(page)
            
            # 판매자 정보 아코디언 찾기 및 클릭
            accordion_clicked = False
            accordion_selectors = [
                "#btnSellerInfo",
                "text=판매자 정보",
                "button:has-text('판매자 정보')",
                "div:has-text('판매자 정보')"
            ]
            
            for selector in accordion_selectors:
                try:
                    accordion = page.locator(selector).first
                    if accordion.count() > 0:
                        accordion.scroll_into_view_if_needed(timeout=3000)
                        time.sleep(0.3)
                        
                        # 이미 열려있는지 확인 (class에 'on'이 있으면 열린 상태)
                        is_open = accordion.evaluate("el => el.classList.contains('on')")
                        if not is_open:
                            accordion.click(timeout=3000)
                            time.sleep(0.5)
                        
                        accordion_clicked = True
                        break
                except:
                    continue
            
            if not accordion_clicked:
                self.log("Warning: Could not find seller info accordion")
                return seller_info
            
            # 판매자 정보 추출
            seller_data = page.evaluate("""
                () => {
                    const data = {};
                    
                    // 1. 테이블 기반 추출 (정밀함)
                    const sellerSection = document.querySelector('.noti_prod_info table, .seller_info_table');
                    if (sellerSection) {
                        const rows = sellerSection.querySelectorAll('tr');
                        rows.forEach(row => {
                            const th = row.querySelector('th');
                            const td = row.querySelector('td');
                            if (th && td) {
                                const label = th.innerText.trim();
                                const value = td.innerText.trim();
                                
                                if (label.includes('상호')) data['판매자명'] = value;
                                else if (label.includes('대표자')) data['대표자명'] = value;
                                else if (label.includes('사업장 소재지') || label.includes('주소')) data['주소'] = value;
                                else if (label.includes('사업자등록번호')) data['사업자등록번호'] = value;
                                else if (label.includes('통신판매업')) data['통신판매업신고'] = value;
                                else if (label.includes('이메일')) data['이메일'] = value;
                                else if (label.includes('연락처') || label.includes('전화')) data['연락처'] = value;
                            }
                        });
                    }
                    
                    // 2. 텍스트 패턴 기반 (폴백 - 테이블이 없거나 레이아웃이 다를 경우)
                    const allText = document.body.innerText;
                    const patterns = {
                        '판매자명': /(?:판매자명|상호)[:\\s]*([^\\n]+)/,
                        '사업자등록번호': /사업자등록번호[:\\s]*([0-9\\-]+)/,
                        '통신판매업신고': /통신판매업[^\\n]*신고[^\\n]*[:\\s]*([^\\n]+)/,
                        '대표자명': /대표자[^\\n]*[:\\s]*([^\\n]+)/,
                        '주소': /(?:주소|소재지|사업장 소재지)[:\\s]*([^\\n]+(?:시|도|구|로)[^\\n]*)/,
                        '연락처': /(?:연락처|전화|Tel)[:\\s]*([0-9\\-]+)/,
                        '이메일': /(?:이메일|E-mail)[:\\s]*([a-zA-Z0-9._%+\\-]+@[a-zA-Z0-9.\\-]+\\.[a-zA-Z]{2,})/
                    };
                    
                    for (const [key, pattern] of Object.entries(patterns)) {
                        if (!data[key]) { // 테이블에서 이미 찾은 정보라면 건너뜀
                            const match = allText.match(pattern);
                            if (match && match[1]) data[key] = match[1].trim();
                        }
                    }
                    
                    return data;
                }
            """)
            
            # 추출된 데이터 병합
            for key, value in seller_data.items():
                if value:
                    seller_info[key] = value
                    
        except Exception as e:
            self.log(f"Error extracting seller info: {e}")
        
        return seller_info
    
    def _close_popups(self, page):
        """팝업 닫기"""
        try:
            # ESC 키로 팝업 닫기
            for _ in range(2):
                try:
                    page.keyboard.press("Escape")
                    time.sleep(0.2)
                except:
                    pass
            
            # JavaScript로 팝업 제거
            try:
                page.evaluate("""
                    () => {
                        const popupSelectors = [
                            '[class*="popup"]',
                            '[class*="modal"]',
                            '[class*="dialog"]',
                            '[class*="overlay"]',
                            '[role="dialog"]',
                            '[class*="reward"]',
                            '[class*="layer"]'
                        ];
                        
                        popupSelectors.forEach(selector => {
                            const elements = document.querySelectorAll(selector);
                            elements.forEach(el => {
                                const zIndex = parseInt(window.getComputedStyle(el).zIndex) || 0;
                                if (zIndex > 100 || el.style.zIndex > 100) {
                                    el.style.display = 'none';
                                    el.remove();
                                }
                            });
                        });
                        
                        document.body.style.overflow = 'auto';
                        document.documentElement.style.overflow = 'auto';
                    }
                """)
            except:
                pass
        except:
            pass
