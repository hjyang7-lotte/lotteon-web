import sys
import os
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

# --- PyInstaller가 .exe로 실행될 때 필요한 경로 설정 코드 ---
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    # .exe로 실행 중일 때
    # 방법 1: 임시 폴더(_MEIPASS) 내의 브라우저 확인
    browsers_path = os.path.join(sys._MEIPASS, 'ms-playwright')
    
    if os.path.exists(browsers_path):
        # Playwright가 브라우저를 찾을 수 있도록 환경 변수를 설정합니다.
        os.environ['PLAYWRIGHT_BROWSERS_PATH'] = browsers_path
    else:
        # 방법 2: 시스템의 기본 Playwright 브라우저 경로 사용
        local_appdata = os.environ.get('LOCALAPPDATA', '')
        system_browsers_path = os.path.join(local_appdata, 'ms-playwright')
        
        if os.path.exists(system_browsers_path):
            os.environ['PLAYWRIGHT_BROWSERS_PATH'] = system_browsers_path
        else:
            # 브라우저 경로를 찾을 수 없을 때
            print(f"경고: Playwright 브라우저를 찾을 수 없습니다.")
            print(f"  - 임시 폴더: {browsers_path}")
            print(f"  - 시스템 경로: {system_browsers_path}")
            print(f"  - Playwright 브라우저가 설치되어 있는지 확인하세요.")
# --- 여기까지 추가 ---

# import tkinter removed for headless environment
import asyncio
from playwright.async_api import async_playwright
import pandas as pd
from datetime import datetime
import threading


class MusinsaCrawler:
    def __init__(self):
        self.stop_flag = False
        self.categories = {
            "전체": "https://www.musinsa.com/main/musinsa/ranking?skip_bf=Y&gf=A&storeCode=musinsa&sectionId=200&contentsId=&categoryCode=000&ageBand=AGE_BAND_ALL",
            "뷰티": "https://www.musinsa.com/main/musinsa/ranking?skip_bf=Y&gf=A&storeCode=musinsa&sectionId=200&contentsId=&categoryCode=104000&ageBand=AGE_BAND_ALL&subPan=product",
            "신발": "https://www.musinsa.com/main/musinsa/ranking?skip_bf=Y&gf=A&storeCode=musinsa&sectionId=200&contentsId=&categoryCode=103000&ageBand=AGE_BAND_ALL&subPan=product",
            "상의": "https://www.musinsa.com/main/musinsa/ranking?skip_bf=Y&gf=A&storeCode=musinsa&sectionId=200&contentsId=&categoryCode=001000&ageBand=AGE_BAND_ALL&subPan=product",
            "아우터": "https://www.musinsa.com/main/musinsa/ranking?skip_bf=Y&gf=A&storeCode=musinsa&sectionId=200&contentsId=&categoryCode=002000&ageBand=AGE_BAND_ALL&subPan=product",
            "바지": "https://www.musinsa.com/main/musinsa/ranking?skip_bf=Y&gf=A&storeCode=musinsa&sectionId=200&contentsId=&categoryCode=003000&ageBand=AGE_BAND_ALL&subPan=product",
            "원피스": "https://www.musinsa.com/main/musinsa/ranking?skip_bf=Y&gf=A&storeCode=musinsa&sectionId=200&contentsId=&categoryCode=100000&ageBand=AGE_BAND_ALL&subPan=product",
            "가방": "https://www.musinsa.com/main/musinsa/ranking?skip_bf=Y&gf=A&storeCode=musinsa&sectionId=200&contentsId=&categoryCode=004000&ageBand=AGE_BAND_ALL&subPan=product",
            "패션소품": "https://www.musinsa.com/main/musinsa/ranking?skip_bf=Y&gf=A&storeCode=musinsa&sectionId=200&contentsId=&categoryCode=101000&ageBand=AGE_BAND_ALL&subPan=product",
            "속옷": "https://www.musinsa.com/main/musinsa/ranking?skip_bf=Y&gf=A&storeCode=musinsa&sectionId=200&contentsId=&categoryCode=026000&ageBand=AGE_BAND_ALL&subPan=product",
            "스포츠": "https://www.musinsa.com/main/musinsa/ranking?skip_bf=Y&gf=A&storeCode=musinsa&sectionId=200&contentsId=&categoryCode=017000&ageBand=AGE_BAND_ALL&subPan=product",
            "키즈": "https://www.musinsa.com/main/musinsa/ranking?skip_bf=Y&gf=A&storeCode=musinsa&sectionId=200&contentsId=&categoryCode=106000&ageBand=AGE_BAND_ALL&subPan=product"
        }
    
    def log(self, message):
        """로그 메시지 출력"""
        if hasattr(self, 'log_callback'):
            self.log_callback(message)
        print(message)
    
    async def get_seller_info(self, page, product_url):
        """상품 페이지에서 판매자 정보 추출 (개선된 로직)"""
        try:
            self.log(f"상품 페이지 접속: {product_url}")
            
            # 1. 페이지 로딩 (속도 우선)
            try:
                await page.goto(product_url, wait_until="domcontentloaded", timeout=30000)
            except Exception as e:
                self.log(f"페이지 로드 타임아웃 (무시): {e}")

            await page.wait_for_timeout(2000)

            # 2. 판매자 정보 섹션 열기 (모든 가능성 시도)
            try:
                # 스크롤을 맨 아래로 내렸다가 다시 올려서 lazy loading 유도
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                await page.wait_for_timeout(500)
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight / 2)")
                await page.wait_for_timeout(500)
                
                # '판매자 정보', '상품 정보 고시' 등이 포함된 버튼/요소 찾아서 클릭
                await page.evaluate("""() => {
                    const searchTerms = ['판매자', '사업자', '정보 고시', '반품'];
                    const elements = document.querySelectorAll('button, a, div[role="button"], h3, h4');
                    
                    for (let el of elements) {
                        const text = el.innerText || '';
                        if (searchTerms.some(term => text.includes(term))) {
                            try { el.click(); } catch(e) {}
                        }
                    }
                }""")
                await page.wait_for_timeout(1500)
            except Exception as e:
                self.log(f"아코디언 클릭 시도 중 오류: {e}")

            # 3. 데이터 추출 (텍스트 기반 범용 탐색)
            seller_info = await page.evaluate("""() => {
                const result = {
                    "상호": "",
                    "사업자번호": "",
                    "연락처": "",
                    "영업소재지": ""
                };
                
                // 페이지 전체 텍스트에서 정규식으로 찾기 (가장 강력함)
                const bodyText = document.body.innerText;
                
                // 정규식 패턴 정의
                const patterns = {
                    "상호": /(?:상호|법인명|업체명)[:\\s]*([^\\n]+)/,
                    "사업자번호": /(?:사업자(?:등록)?번호)[:\\s]*([0-9\\-]+)/,
                    "연락처": /(?:연락처|전화번호|문의)[:\\s]*([0-9\\-]+)/,
                    "영업소재지": /(?:소재지|주소)[:\\s]*([^\\n]+(?:시|도|구|동)[^\\n]*)/
                };

                // 1차 시도: 테이블이나 DL 구조에서 정확히 매칭
                const tables = document.querySelectorAll('table, dl');
                tables.forEach(table => {
                    const text = table.innerText;
                    if (text.includes('상호') || text.includes('사업자')) {
                        // 행 단위 분석
                        const rows = table.querySelectorAll('tr, div, li');
                        rows.forEach(row => {
                            const rowText = row.innerText;
                            if (rowText.includes('상호')) result["상호"] = rowText.split('상호')[1]?.trim() || result["상호"];
                            if (rowText.includes('대표')) result["상호"] = rowText.split('대표')[1]?.trim() || result["상호"]; // 대표자도 상호에 포함 고려
                            if (rowText.includes('사업자')) result["사업자번호"] = rowText.replace(/[^0-9\\-]/g, '');
                            if (rowText.includes('전화') || rowText.includes('연락처')) result["연락처"] = rowText.replace(/[^0-9\\-]/g, '');
                            if (rowText.includes('소재지') || rowText.includes('주소')) {
                                result["영업소재지"] = rowText.split(/(?:소재지|주소)/)[1]?.trim() || result["영업소재지"];
                            }
                        });
                    }
                });

                // 2차 시도: 정규식으로 보완
                for (const [key, regex] of Object.entries(patterns)) {
                    if (!result[key] || result[key].length < 2) {
                        const match = bodyText.match(regex);
                        if (match && match[1]) {
                            result[key] = match[1].trim();
                        }
                    }
                }
                
                return result;
            }""")
            
            # 정제
            if seller_info:
                # 상호가 비어있고 대표자가 있다면 대표자를 상호로 (임시)
                if not seller_info["상호"] and "대표자" in seller_info: # JS에서 대표자를 상호로 넣긴 했음
                    pass 
                
                # 결과 정리
                final_info = {
                    "상호": seller_info.get("상호", "").strip()[:50],
                    "사업자번호": seller_info.get("사업자번호", "").strip(),
                    "연락처": seller_info.get("연락처", "").strip(),
                    "영업소재지": seller_info.get("영업소재지", "").strip()[:100]
                }
                
                self.log(f"판매자 정보 추출 완료: {final_info}")
                return final_info
            
            return { "상호": "", "사업자번호": "", "연락처": "", "영업소재지": "" }

        except Exception as e:
            self.log(f"판매자 정보 로직 오류: {str(e)}")
            return { "상호": "", "사업자번호": "", "연락처": "", "영업소재지": "" }
    
    async def crawl_products(self, category, url, num_products, progress_callback=None):
        """상품 크롤링 실행"""
        products = []
        
        async with async_playwright() as p:
            self.log("브라우저 실행 중...")
            browser = await p.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-dev-shm-usage"]
            )
            # 타임아웃 증가
            page = await browser.new_page()
            page.set_default_timeout(60000)
            
            try:
                self.log(f"{category} 카테고리 페이지 로딩 중...")
                # 페이지 로드 전략 간소화: domcontentloaded만 기다리고 바로 시작 (속도 향상)
                try:
                    await page.goto(url, wait_until="domcontentloaded", timeout=45000)
                except Exception as e:
                    self.log(f"초기 로딩 타임아웃 (계속 진행): {e}")

                # 상품이 로드될 때까지 잠시 대기
                try:
                    await page.wait_for_selector('a.gtm-select-item', timeout=10000)
                except:
                    self.log("상품 목록 선택자 대기 실패, 스크롤 시도")

                # 스크롤 최적화
                self.log("상품 목록 로딩 중...")
                for i in range(10):  # 최대 횟수 줄임
                    if self.stop_flag:
                        break
                    
                    # 현재 개수 체크 - 충분하면 즉시 중단 (속도 핵심)
                    current_items = await page.query_selector_all('a.gtm-select-item')
                    current_count = len(current_items)
                    self.log(f"스크롤 {i+1}/10 - 현재 발견된 상품: {current_count}개 (목표: {num_products}개)")
                    
                    if current_count >= num_products:
                        self.log("충분한 상품을 찾았습니다. 스크롤 중단.")
                        break

                    await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                    await page.wait_for_timeout(1000) # 대기 시간 단축
                    
                    # 스크롤 후 상품 개수 확인
                    new_count = len(await page.query_selector_all('a.gtm-select-item'))
                    if new_count == current_count and i > 5:
                        self.log("더 이상 새로운 상품이 로드되지 않습니다.")
                        break
                
                # 페이지가 완전히 로드될 때까지 추가 대기
                await page.wait_for_timeout(5000)
                
                # JavaScript 실행 완료 대기 - 상품이 동적으로 로드될 수 있음
                self.log("JavaScript 실행 완료 대기 중...")
                try:
                    # 페이지의 JavaScript가 완료될 때까지 대기
                    await page.evaluate('''() => {
                        return new Promise((resolve) => {
                            if (document.readyState === 'complete') {
                                setTimeout(resolve, 3000);
                            } else {
                                window.addEventListener('load', () => {
                                    setTimeout(resolve, 3000);
                                });
                            }
                        });
                    }''')
                    self.log("JavaScript 실행 완료")
                except Exception as e:
                    self.log(f"JavaScript 대기 중 오류 (무시): {str(e)}")
                
                # 추가 대기
                await page.wait_for_timeout(5000)
                
                # 실제 페이지에 상품이 있는지 확인
                self.log("페이지 내용 확인 중...")
                page_text = await page.evaluate('() => document.body.innerText')
                if '상품' in page_text or 'product' in page_text.lower():
                    self.log("페이지에 상품 관련 텍스트 발견")
                else:
                    self.log("경고: 페이지에 상품 관련 텍스트를 찾지 못했습니다")
                
                # 상품 정보 추출 - 여러 셀렉터 시도
                product_items = []
                
                # 방법 1: 기존 셀렉터
                product_items = await page.query_selector_all('div.UIProductColumn__InfoItem-sc-1t5ihy5-7')
                self.log(f"셀렉터 1 결과: {len(product_items)}개 상품 발견")
                
                # 방법 2: 더 일반적인 셀렉터 시도
                if len(product_items) == 0:
                    product_items = await page.query_selector_all('div[class*="UIProductColumn"]')
                    self.log(f"셀렉터 2 결과: {len(product_items)}개 상품 발견")
                
                # 방법 3: 상품 링크로 찾기 - 링크를 기준으로 부모 컨테이너 찾기
                if len(product_items) == 0:
                    product_links = await page.query_selector_all('a.gtm-select-item')
                    self.log(f"셀렉터 3 (링크) 결과: {len(product_links)}개 상품 링크 발견")
                    # 링크의 부모 요소 찾기
                    if product_links:
                        product_items = []
                        for link in product_links:
                            try:
                                # JavaScript로 부모 컨테이너의 선택자 찾기
                                parent_selector = await link.evaluate('''el => {
                                    let current = el.parentElement;
                                    let depth = 0;
                                    while (current && depth < 10) {
                                        if (current.classList && current.classList.toString().includes('UIProductColumn')) {
                                            // 클래스명으로 선택자 생성
                                            const classes = Array.from(current.classList).join('.');
                                            return `div.${classes}`;
                                        }
                                        current = current.parentElement;
                                        depth++;
                                    }
                                    return null;
                                }''')
                                if parent_selector:
                                    # 찾은 선택자로 요소 찾기
                                    parent_elements = await page.query_selector_all(parent_selector)
                                    # 같은 링크를 포함하는 부모 찾기
                                    for parent in parent_elements:
                                        try:
                                            link_in_parent = await parent.query_selector('a.gtm-select-item')
                                            if link_in_parent:
                                                link_href = await link.get_attribute('href')
                                                parent_link_href = await link_in_parent.get_attribute('href')
                                                if link_href == parent_link_href:
                                                    product_items.append(parent)
                                                    break
                                        except:
                                            continue
                            except Exception as e:
                                self.log(f"부모 요소 찾기 오류: {str(e)}")
                                continue
                        # 중복 제거
                        seen = set()
                        unique_items = []
                        for item in product_items:
                            try:
                                item_id = await item.evaluate('el => el.outerHTML.substring(0, 100)')
                                if item_id not in seen:
                                    seen.add(item_id)
                                    unique_items.append(item)
                            except:
                                unique_items.append(item)
                        product_items = unique_items
                        self.log(f"부모 요소 찾기 결과: {len(product_items)}개 상품 발견")
                
                # 방법 4: 랭킹 페이지의 일반적인 상품 컨테이너 찾기
                if len(product_items) == 0:
                    # 페이지 구조 디버깅 - 더 자세한 정보 수집
                    self.log("페이지 구조 분석 중...")
                    all_divs = await page.query_selector_all('div')
                    self.log(f"전체 div 개수: {len(all_divs)}")
                    
                    # 다양한 링크 패턴 확인
                    all_links = await page.query_selector_all('a[href*="/products/"]')
                    self.log(f"/products/ 링크 개수: {len(all_links)}")
                    
                    # 다른 링크 패턴도 확인
                    ranking_links = await page.query_selector_all('a[href*="ranking"]')
                    self.log(f"ranking 링크 개수: {len(ranking_links)}")
                    
                    # 클래스명에 product가 포함된 요소 찾기 (대소문자 구분)
                    product_divs = await page.query_selector_all('[class*="product"], [class*="Product"]')
                    self.log(f"product/Product 클래스 포함 요소: {len(product_divs)}개")
                    
                    # gtm 관련 요소 찾기
                    gtm_elements = await page.query_selector_all('[class*="gtm"]')
                    self.log(f"gtm 클래스 포함 요소: {len(gtm_elements)}개")
                    
                    # 실제 페이지의 모든 링크 클래스 확인
                    all_a_tags = await page.query_selector_all('a')
                    self.log(f"전체 링크(a 태그) 개수: {len(all_a_tags)}")
                    
                    # 링크의 클래스명 샘플 수집
                    link_classes = set()
                    for link in all_a_tags[:20]:
                        try:
                            classes = await link.get_attribute('class')
                            if classes:
                                link_classes.add(classes)
                        except:
                            pass
                    if link_classes:
                        self.log(f"링크 클래스 샘플 (최대 10개): {list(link_classes)[:10]}")
                    
                    # 실제 링크 URL 샘플 확인
                    if all_links:
                        sample_urls = []
                        for link in all_links[:5]:
                            try:
                                href = await link.get_attribute('href')
                                if href:
                                    sample_urls.append(href)
                            except:
                                pass
                        self.log(f"샘플 링크 URL: {sample_urls}")
                    else:
                        # /products/ 링크가 없으면 다른 패턴 확인
                        sample_hrefs = []
                        for link in all_a_tags[:10]:
                            try:
                                href = await link.get_attribute('href')
                                if href and ('product' in href.lower() or 'item' in href.lower()):
                                    sample_hrefs.append(href[:100])  # 처음 100자만
                            except:
                                pass
                        if sample_hrefs:
                            self.log(f"상품 관련 링크 샘플: {sample_hrefs}")
                    
                    # 페이지의 실제 HTML 구조 일부 확인 (디버깅용)
                    try:
                        html_sample = await page.evaluate('''() => {
                            const body = document.body;
                            if (!body) return "body 없음";
                            const html = body.innerHTML.substring(0, 2000);
                            return html;
                        }''')
                        # HTML에서 상품 관련 키워드 찾기
                        if 'product' in html_sample.lower() or '상품' in html_sample:
                            self.log("HTML에 상품 관련 내용 발견")
                        else:
                            self.log("경고: HTML에 상품 관련 내용을 찾지 못했습니다")
                    except Exception as e:
                        self.log(f"HTML 샘플 확인 중 오류: {str(e)}")
                    
                    # gtm-select-item 클래스를 가진 링크의 부모 찾기
                    product_links = await page.query_selector_all('a.gtm-select-item')
                    if product_links:
                        self.log(f"gtm-select-item 링크 {len(product_links)}개 발견")
                        # 각 링크를 직접 사용 (링크 자체에서 정보 추출 가능)
                        # 링크 주변의 정보를 추출할 수 있도록 링크를 기준으로 작업
                        product_items = []
                        for link in product_links[:num_products]:
                            try:
                                # 링크의 부모 컨테이너 선택자 찾기
                                container_selector = await link.evaluate('''el => {
                                    let current = el.parentElement;
                                    let depth = 0;
                                    while (current && depth < 15) {
                                        const classList = current.classList ? current.classList.toString() : '';
                                        if (classList.includes('UIProductColumn') || 
                                            classList.includes('ProductColumn') ||
                                            (classList.includes('product') && current.tagName === 'DIV')) {
                                            const classes = Array.from(current.classList).filter(c => c.length > 0);
                                            if (classes.length > 0) {
                                                return `div.${classes.join('.')}`;
                                            }
                                            return `div[class*="${classList.substring(0, 20)}"]`;
                                        }
                                        current = current.parentElement;
                                        depth++;
                                    }
                                    return null;
                                }''')
                                if container_selector:
                                    # 선택자로 요소 찾기
                                    try:
                                        containers = await page.query_selector_all(container_selector)
                                        # 링크를 포함하는 컨테이너 찾기
                                        for container in containers:
                                            try:
                                                link_in_container = await container.query_selector('a.gtm-select-item')
                                                if link_in_container:
                                                    link_href = await link.get_attribute('href')
                                                    container_link_href = await link_in_container.get_attribute('href')
                                                    if link_href and container_link_href and link_href == container_link_href:
                                                        product_items.append(container)
                                                        break
                                            except:
                                                continue
                                    except:
                                        pass
                            except Exception as e:
                                self.log(f"컨테이너 찾기 오류: {str(e)}")
                                continue
                        # 중복 제거
                        seen_urls = set()
                        unique_items = []
                        for item in product_items:
                            try:
                                link_elem = await item.query_selector('a.gtm-select-item')
                                if link_elem:
                                    url = await link_elem.get_attribute('href')
                                    if url and url not in seen_urls:
                                        seen_urls.add(url)
                                        unique_items.append(item)
                                else:
                                    unique_items.append(item)
                            except:
                                unique_items.append(item)
                        product_items = unique_items
                        self.log(f"부모 컨테이너 찾기 결과: {len(product_items)}개 상품 발견")
                
                # 최종 방법: 링크를 직접 사용 (컨테이너를 찾지 못한 경우)
                if len(product_items) == 0:
                    # gtm-select-item 링크 시도
                    product_links = await page.query_selector_all('a.gtm-select-item')
                    if product_links:
                        self.log(f"최종 방법 1: gtm-select-item 링크 사용 ({len(product_links)}개 링크 발견)")
                        product_items = product_links
                    else:
                        # /products/ 링크 직접 사용
                        product_links = await page.query_selector_all('a[href*="/products/"]')
                        if product_links:
                            self.log(f"최종 방법 2: /products/ 링크 사용 ({len(product_links)}개 링크 발견)")
                            # 중복 제거
                            seen_urls = set()
                            unique_links = []
                            for link in product_links:
                                try:
                                    href = await link.get_attribute('href')
                                    if href and href not in seen_urls:
                                        seen_urls.add(href)
                                        unique_links.append(link)
                                except:
                                    continue
                            product_items = unique_links
                            self.log(f"중복 제거 후: {len(product_items)}개 링크")
                
                total_items = min(len(product_items), num_products)
                self.log(f"총 {total_items}개 상품 크롤링 시작...")
                
                if total_items == 0:
                    self.log("경고: 상품을 찾을 수 없습니다. 페이지 구조를 확인하세요.")
                    # 페이지 스크린샷 저장 (디버깅용)
                    try:
                        await page.screenshot(path="debug_page.png")
                        self.log("디버깅용 스크린샷 저장: debug_page.png")
                    except:
                        pass
                    return products  # 빈 리스트 반환
                
                # 1단계: 먼저 모든 상품의 기본 정보 수집
                product_urls = []
                basic_info_list = []
                
                for idx, item in enumerate(product_items[:num_products]):
                    if self.stop_flag:
                        self.log("크롤링 중지됨")
                        break
                    
                    try:
                        # 랭킹
                        rank = idx + 1
                        
                        # item이 링크인지 컨테이너인지 확인
                        item_tag = await item.evaluate('el => el.tagName')
                        is_link = (item_tag == 'A')
                        
                        # 브랜드명 - 여러 방법 시도
                        brand = ""
                        if is_link:
                            # 링크인 경우: 부모에서 브랜드 찾기
                            try:
                                # 링크의 부모 컨테이너에서 브랜드 찾기
                                brand_text = await item.evaluate('''el => {
                                    let current = el.parentElement;
                                    let depth = 0;
                                    while (current && depth < 10) {
                                        // 브랜드 링크 찾기
                                        const brandLink = current.querySelector('a.gtm-click-brand');
                                        if (brandLink) {
                                            const brandP = brandLink.querySelector('p');
                                            if (brandP) return brandP.textContent.trim();
                                            return brandLink.textContent.trim();
                                        }
                                        // 브랜드 텍스트 찾기
                                        const brandP = current.querySelector('p[class*="brand"]');
                                        if (brandP) return brandP.textContent.trim();
                                        const brandSpan = current.querySelector('span[class*="brand"]');
                                        if (brandSpan) return brandSpan.textContent.trim();
                                        current = current.parentElement;
                                        depth++;
                                    }
                                    return "";
                                }''')
                                if brand_text:
                                    brand = brand_text
                            except:
                                pass
                        else:
                            # 컨테이너인 경우: 기존 방법
                            brand_selectors = [
                                'a.gtm-click-brand p',
                                'a.gtm-click-brand',
                                'p[class*="brand"]',
                                'span[class*="brand"]'
                            ]
                            for selector in brand_selectors:
                                try:
                                    brand_element = await item.query_selector(selector)
                                    if brand_element:
                                        brand = await brand_element.inner_text()
                                        if brand.strip():
                                            break
                                except:
                                    continue
                        
                        # 상품명 및 URL - 여러 방법 시도
                        product_name = ""
                        product_url = ""
                        
                        if is_link:
                            # 링크인 경우: 직접 사용
                            product_element = item
                            # 상품명 추출
                            name_selectors = ['p', 'span', 'div']
                            for name_sel in name_selectors:
                                try:
                                    name_element = await product_element.query_selector(name_sel)
                                    if name_element:
                                        product_name = await name_element.inner_text()
                                        if product_name.strip():
                                            break
                                except:
                                    continue
                            
                            # URL 추출
                            product_url = await product_element.get_attribute('href')
                            if product_url and not product_url.startswith('http'):
                                product_url = f"https://www.musinsa.com{product_url}"
                        else:
                            # 컨테이너인 경우: 기존 방법
                            product_selectors = [
                                'a.gtm-select-item',
                                'a[href*="/products/"]',
                                'a[class*="product"]'
                            ]
                            for selector in product_selectors:
                                try:
                                    product_element = await item.query_selector(selector)
                                    if product_element:
                                        # 상품명 추출
                                        name_selectors = ['p', 'span', 'div']
                                        for name_sel in name_selectors:
                                            try:
                                                name_element = await product_element.query_selector(name_sel)
                                                if name_element:
                                                    product_name = await name_element.inner_text()
                                                    if product_name.strip():
                                                        break
                                            except:
                                                continue
                                        
                                        # URL 추출
                                        product_url = await product_element.get_attribute('href')
                                        if product_url:
                                            if not product_url.startswith('http'):
                                                product_url = f"https://www.musinsa.com{product_url}"
                                            break
                                except:
                                    continue
                        
                        # 가격 정보 - 여러 방법 시도
                        discount_rate = ""
                        price = ""
                        
                        if is_link:
                            # 링크인 경우: 부모에서 가격 찾기
                            try:
                                # JavaScript로 부모에서 가격 정보 추출
                                price_info = await item.evaluate('''el => {
                                    let current = el.parentElement;
                                    let depth = 0;
                                    while (current && depth < 10) {
                                        // 가격 컨테이너 찾기
                                        const priceDiv = current.querySelector('div.UIProductColumn__Price-sc-1t5ihy5-10') ||
                                                         current.querySelector('div[class*="Price"]') ||
                                                         current.querySelector('span[class*="price"]') ||
                                                         current.querySelector('p[class*="price"]');
                                        if (priceDiv) {
                                            // 할인율
                                            const discountSpan = priceDiv.querySelector('span.text-red') ||
                                                                 priceDiv.querySelector('span[class*="red"]') ||
                                                                 priceDiv.querySelector('span[class*="discount"]');
                                            const discount = discountSpan ? discountSpan.textContent.trim() : "";
                                            
                                            // 가격
                                            const priceSpan = priceDiv.querySelector('span.text-black') ||
                                                              priceDiv.querySelector('span[class*="black"]') ||
                                                              priceDiv.querySelector('span') ||
                                                              priceDiv.querySelector('p');
                                            let price = priceSpan ? priceSpan.textContent.trim() : "";
                                            
                                            // 숫자가 포함된 경우만 가격으로 인식
                                            if (price && !/[0-9]/.test(price)) {
                                                // span이나 p의 모든 자식 요소 확인
                                                const allSpans = priceDiv.querySelectorAll('span, p');
                                                for (let sp of allSpans) {
                                                    const text = sp.textContent.trim();
                                                    if (/[0-9]/.test(text)) {
                                                        price = text;
                                                        break;
                                                    }
                                                }
                                            }
                                            
                                            return { discount: discount, price: price };
                                        }
                                        current = current.parentElement;
                                        depth++;
                                    }
                                    return { discount: "", price: "" };
                                }''')
                                if price_info:
                                    discount_rate = price_info.get('discount', '')
                                    price = price_info.get('price', '')
                            except Exception as e:
                                self.log(f"가격 정보 추출 오류: {str(e)}")
                                pass
                        else:
                            # 컨테이너인 경우: 기존 방법
                            price_selectors = [
                                'div.UIProductColumn__Price-sc-1t5ihy5-10',
                                'div[class*="Price"]',
                                'span[class*="price"]',
                                'p[class*="price"]'
                            ]
                            for selector in price_selectors:
                                try:
                                    price_element = await item.query_selector(selector)
                                    if price_element:
                                        # 할인율
                                        discount_selectors = ['span.text-red', 'span[class*="red"]', 'span[class*="discount"]']
                                        for disc_sel in discount_selectors:
                                            try:
                                                discount_element = await price_element.query_selector(disc_sel)
                                                if discount_element:
                                                    discount_rate = await discount_element.inner_text()
                                                    if discount_rate.strip():
                                                        break
                                            except:
                                                continue
                                        
                                        # 가격
                                        price_selectors_inner = ['span.text-black', 'span[class*="black"]', 'span', 'p']
                                        for price_sel in price_selectors_inner:
                                            try:
                                                price_element_text = await price_element.query_selector(price_sel)
                                                if price_element_text:
                                                    price_text = await price_element_text.inner_text()
                                                    # 숫자만 포함된 경우만 가격으로 인식
                                                    if any(c.isdigit() for c in price_text):
                                                        price = price_text
                                                        if price.strip():
                                                            break
                                            except:
                                                continue
                                        
                                        if price.strip():
                                            break
                                except:
                                    continue
                        
                        # 기본 정보 저장
                        basic_info_list.append({
                            "카테고리": category,
                            "랭킹": rank,
                            "브랜드": brand.strip(),
                            "상품명": product_name.strip(),
                            "할인율": discount_rate.strip(),
                            "가격": price.strip(),
                            "상품URL": product_url
                        })
                        product_urls.append(product_url)
                        
                        self.log(f"상품 {idx + 1} 정보 수집 완료: {brand.strip()} - {product_name.strip()}")
                        
                    except Exception as e:
                        self.log(f"상품 {idx + 1} 기본 정보 수집 중 오류: {str(e)}")
                        import traceback
                        self.log(f"상세 오류: {traceback.format_exc()}")
                        # 오류 발생 시 빈 정보 추가
                        basic_info_list.append({
                            "카테고리": category,
                            "랭킹": idx + 1,
                            "브랜드": "",
                            "상품명": "",
                            "할인율": "",
                            "가격": "",
                            "상품URL": ""
                        })
                        product_urls.append("")
                        continue
                
                # 2단계: 각 상품의 판매자 정보 수집
                self.log("판매자 정보 수집 시작...")
                for idx, (basic_info, product_url) in enumerate(zip(basic_info_list, product_urls)):
                    if self.stop_flag:
                        self.log("크롤링 중지됨")
                        break
                    
                    if not product_url:
                        # URL이 없는 경우 빈 판매자 정보 추가
                        basic_info.update({
                            "상호": "",
                            "사업자번호": "",
                            "연락처": "",
                            "영업소재지": ""
                        })
                        products.append(basic_info)
                        if progress_callback:
                            progress_callback(idx + 1, total_items)
                        continue
                    
                    try:
                        self.log(f"[{idx + 1}/{total_items}] {basic_info['브랜드']} - {basic_info['상품명']} 판매자 정보 수집 중...")
                        seller_info = await self.get_seller_info(page, product_url)
                        
                        basic_info.update({
                            "상호": seller_info["상호"],
                            "사업자번호": seller_info["사업자번호"],
                            "연락처": seller_info["연락처"],
                            "영업소재지": seller_info["영업소재지"]
                        })
                        
                        products.append(basic_info)
                        
                        if progress_callback:
                            progress_callback(idx + 1, total_items)
                        
                    except Exception as e:
                        self.log(f"상품 {idx + 1} 판매자 정보 수집 중 오류: {str(e)}")
                        # 오류 발생 시 빈 판매자 정보 추가
                        basic_info.update({
                            "상호": "",
                            "사업자번호": "",
                            "연락처": "",
                            "영업소재지": ""
                        })
                        products.append(basic_info)
                        if progress_callback:
                            progress_callback(idx + 1, total_items)
                        continue
                
            except Exception as e:
                self.log(f"크롤링 중 오류 발생: {str(e)}")
            
            finally:
                await browser.close()
        
        return products
    
    def save_to_excel(self, products, category, output_dir=None):
        """엑셀 파일로 저장"""
        if not products:
            self.log("저장할 데이터가 없습니다.")
            return None
        
        df = pd.DataFrame(products)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"musinsa_{category}_{timestamp}.xlsx"
        
        # output_dir이 없으면 기본적으로 상위 폴더의 results에 저장 시도
        if output_dir is None:
            # 현재 스크립트 위치: .../lotteon md/musinsa best new/
            # 목표: .../lotteon md/results/
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            output_dir = os.path.join(base_dir, "results")

        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except:
                pass

        filepath = os.path.join(output_dir, filename)
        
        df.to_excel(filepath, index=False, engine='openpyxl')
        self.log(f"파일 저장 완료: {filepath}")
        return filepath


class MusinsaCrawlerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("무신사 베스트 상품 크롤러")
        self.root.geometry("700x600")
        self.root.resizable(False, False)
        
        # 창을 화면 중앙에 배치
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
        # 창을 맨 앞으로 가져오기
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after_idle(lambda: self.root.attributes('-topmost', False))
        
        self.crawler = MusinsaCrawler()
        self.crawler.log_callback = self.add_log
        
        self.create_widgets()
        
        # 창이 제대로 표시되도록 강제 업데이트
        self.root.update_idletasks()
        self.root.deiconify()  # 창이 숨겨져 있으면 표시
        self.root.lift()  # 창을 위로 올리기
        self.root.focus_force()  # 포커스 강제 설정
        
        # Windows에서 창을 맨 앞으로 가져오기
        try:
            self.root.attributes('-topmost', True)
            self.root.update()
            self.root.after(100, lambda: self.root.attributes('-topmost', False))
        except:
            pass
        
    def create_widgets(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 카테고리 선택
        ttk.Label(main_frame, text="카테고리:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.category_var = tk.StringVar()
        category_combo = ttk.Combobox(main_frame, textvariable=self.category_var, 
                                       values=list(self.crawler.categories.keys()),
                                       state='readonly', width=30)
        category_combo.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        category_combo.current(0)
        
        # 상품 개수 입력
        ttk.Label(main_frame, text="크롤링 개수:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.num_products_var = tk.StringVar(value="10")
        num_entry = ttk.Entry(main_frame, textvariable=self.num_products_var, width=32)
        num_entry.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        
        # 진행 상황
        ttk.Label(main_frame, text="진행 상황:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.progress_var = tk.StringVar(value="대기 중...")
        progress_label = ttk.Label(main_frame, textvariable=self.progress_var, foreground='blue')
        progress_label.grid(row=2, column=1, sticky=tk.W, pady=5, padx=5)
        
        # 프로그레스 바
        self.progress_bar = ttk.Progressbar(main_frame, mode='determinate', length=400)
        self.progress_bar.grid(row=3, column=0, columnspan=2, pady=10, padx=5)
        
        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        self.start_button = ttk.Button(button_frame, text="실행", command=self.start_crawling, width=15)
        self.start_button.grid(row=0, column=0, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="중지", command=self.stop_crawling, 
                                       state='disabled', width=15)
        self.stop_button.grid(row=0, column=1, padx=5)
        
        # 로그 영역
        ttk.Label(main_frame, text="로그:", font=('Arial', 10, 'bold')).grid(row=5, column=0, sticky=tk.W, pady=5)
        
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, width=80, height=20, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
    def add_log(self, message):
        """로그 메시지 추가"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def update_progress(self, current, total):
        """진행률 업데이트"""
        percentage = (current / total) * 100
        self.progress_bar['value'] = percentage
        self.progress_var.set(f"{current}/{total} 완료 ({percentage:.1f}%)")
        self.root.update()
    
    def start_crawling(self):
        """크롤링 시작"""
        try:
            num_products = int(self.num_products_var.get())
            if num_products <= 0:
                messagebox.showerror("오류", "1개 이상의 상품을 입력하세요.")
                return
        except ValueError:
            messagebox.showerror("오류", "올바른 숫자를 입력하세요.")
            return
        
        category = self.category_var.get()
        url = self.crawler.categories[category]
        
        # UI 상태 변경
        self.start_button['state'] = 'disabled'
        self.stop_button['state'] = 'normal'
        self.crawler.stop_flag = False
        self.progress_bar['value'] = 0
        self.progress_var.set("크롤링 시작...")
        self.log_text.delete(1.0, tk.END)
        
        # 별도 스레드에서 크롤링 실행
        thread = threading.Thread(target=self.run_crawling, args=(category, url, num_products))
        thread.daemon = True
        thread.start()
    
    def run_crawling(self, category, url, num_products):
        """크롤링 실행 (비동기)"""
        try:
            # 새로운 이벤트 루프 생성
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            
            products = loop.run_until_complete(
                self.crawler.crawl_products(category, url, num_products, self.update_progress)
            )
            
            if products and not self.crawler.stop_flag:
                filename = self.crawler.save_to_excel(products, category)
                self.add_log(f"✓ 크롤링 완료! 총 {len(products)}개 상품")
                self.add_log(f"✓ 저장 위치: {os.path.abspath(filename)}")
                self.progress_var.set(f"완료: {len(products)}개 상품 저장됨")
                messagebox.showinfo("완료", f"크롤링이 완료되었습니다!\n파일: {filename}")
            elif self.crawler.stop_flag:
                self.add_log("크롤링이 중지되었습니다.")
                self.progress_var.set("중지됨")
            else:
                self.add_log("크롤링된 상품이 없습니다.")
                self.progress_var.set("실패")
                
        except Exception as e:
            self.add_log(f"오류 발생: {str(e)}")
            self.progress_var.set("오류 발생")
            messagebox.showerror("오류", f"크롤링 중 오류가 발생했습니다:\n{str(e)}")
        
        finally:
            # UI 상태 복원
            self.start_button['state'] = 'normal'
            self.stop_button['state'] = 'disabled'
    
    def stop_crawling(self):
        """크롤링 중지"""
        self.crawler.stop_flag = True
        self.add_log("크롤링 중지 요청됨...")
        self.stop_button['state'] = 'disabled'


if __name__ == "__main__":
    import sys
    
    # Windows에서 GUI가 제대로 표시되도록 설정
    if sys.platform == 'win32':
        import ctypes
        # 콘솔 창 숨기기 (선택사항)
        # ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    
    root = tk.Tk()
    app = MusinsaCrawlerGUI(root)
    
    # 창이 확실히 표시되도록 강제 업데이트
    root.update()
    root.deiconify()
    root.lift()
    root.focus_force()
    
    # Windows에서 창을 맨 앞으로 가져오기
    try:
        root.attributes('-topmost', True)
        root.update()
        root.after(200, lambda: root.attributes('-topmost', False))
    except:
        pass
    
    # 창이 표시되었는지 확인
    print("GUI 창이 표시되었습니다. 창이 보이지 않으면 작업 표시줄을 확인하세요.")
    
    root.mainloop()

