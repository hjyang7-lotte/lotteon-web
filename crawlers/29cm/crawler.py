"""
29cm 크롤러 배포용 메인 파일
GUI 기반 웹 크롤링 애플리케이션
"""

try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
except ImportError:
    # Headless Environment Mock
    class MockTk:
        def __getattr__(self, name): return self
        def __call__(self, *args, **kwargs): return self
        W = E = N = S = VERTICAL = HORIZONTAL = RIGHT = LEFT = BOTH = Y = END = 'mock'
        StringVar = lambda *a, **k: MockVar()
        
    class MockVar:
        def __init__(self, value=None): self.value = value
        def get(self): return self.value
        def set(self, v): self.value = v

    tk = MockTk()
    ttk = MockTk()
    messagebox = MockTk()
    filedialog = MockTk()
import asyncio
from playwright.async_api import async_playwright
import pandas as pd
from datetime import datetime
import os
import sys
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass


class CrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("29cm 크롤러")
        self.root.geometry("600x400")
        
        # 변수 초기화
        try:
            self.search_keyword = tk.StringVar(value="여성가방")
        except (RuntimeError, AttributeError):
            # GUI 없이 실행될 경우 (Wrapper 모드)
            class MockVar:
                def __init__(self, value=''): self.val = value
                def get(self): return self.val
                def set(self, value): self.val = value
            self.search_keyword = MockVar("여성가방")
            
        self.is_crawling = False
        
        try:
            self.setup_ui()
        except Exception:
            # Wrapper 모드에서는 UI 셋업 실패해도 무시
            pass
        
    def setup_ui(self):
        """UI 구성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 검색 키워드 입력
        ttk.Label(main_frame, text="검색 키워드:").grid(row=0, column=0, sticky=tk.W, pady=5)
        keyword_entry = ttk.Entry(main_frame, textvariable=self.search_keyword, width=40)
        keyword_entry.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # 크롤링 시작 버튼
        self.start_button = ttk.Button(main_frame, text="크롤링 시작", command=self.start_crawling)
        self.start_button.grid(row=1, column=0, columnspan=3, pady=10)
        
        # 진행 상태 표시
        self.progress_var = tk.StringVar(value="대기 중...")
        progress_label = ttk.Label(main_frame, textvariable=self.progress_var)
        progress_label.grid(row=2, column=0, columnspan=3, pady=5)
        
        # 진행 바
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # 로그 텍스트 영역
        log_frame = ttk.LabelFrame(main_frame, text="로그", padding="5")
        log_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        self.log_text = tk.Text(log_frame, height=10, width=70)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
    def log(self, message):
        """로그 메시지 추가"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def start_crawling(self):
        """크롤링 시작"""
        if self.is_crawling:
            messagebox.showwarning("경고", "이미 크롤링이 진행 중입니다.")
            return
            
        keyword = self.search_keyword.get().strip()
        if not keyword:
            messagebox.showerror("오류", "검색 키워드를 입력해주세요.")
            return
            
        self.is_crawling = True
        self.start_button.config(state='disabled')
        self.progress_bar.start()
        self.progress_var.set("크롤링 진행 중...")
        
        # 비동기 크롤링 실행
        asyncio.run(self.crawl_29cm(keyword))
        
    async def crawl_29cm(self, keyword, category=None, count=50):
        """29cm 크롤링 메인 함수 (상세 페이지 판매자 정보 수집 기능 추가)"""
        try:
            # 카테고리 URL 매핑
            category_urls = {
                "전체": "https://home.29cm.co.kr/best-products?period=HOURLY&ranking=POPULARITY&gender=F&age=30",
                "여성의류": "https://home.29cm.co.kr/best-products?period=HOURLY&ranking=POPULARITY&gender=F&age=30&categoryLargeCode=268100100",
                "여성가방": "https://home.29cm.co.kr/best-products?period=HOURLY&ranking=POPULARITY&gender=F&age=30&categoryLargeCode=269100100",
                "여성슈즈": "https://home.29cm.co.kr/best-products?period=HOURLY&ranking=POPULARITY&gender=F&age=30&categoryLargeCode=270100100",
                "악세서리": "https://home.29cm.co.kr/best-products?period=HOURLY&ranking=POPULARITY&gender=F&age=30&categoryLargeCode=271100100",
                "주얼리": "https://home.29cm.co.kr/best-products?period=HOURLY&ranking=POPULARITY&gender=F&age=30&categoryLargeCode=305100100",
                "뷰티": "https://home.29cm.co.kr/best-products?period=HOURLY&ranking=POPULARITY&gender=F&age=30&categoryLargeCode=266100100",
                "레저": "https://home.29cm.co.kr/best-products?period=HOURLY&ranking=POPULARITY&gender=F&age=30&categoryLargeCode=286100100",
                "키즈": "https://home.29cm.co.kr/best-products?period=HOURLY&ranking=POPULARITY&gender=F&age=30&categoryLargeCode=290100100",
                "남성의류": "https://home.29cm.co.kr/best-products?period=HOURLY&ranking=POPULARITY&gender=F&age=30&categoryLargeCode=272100100",
                "남성가방": "https://home.29cm.co.kr/best-products?period=HOURLY&ranking=POPULARITY&gender=F&age=30&categoryLargeCode=273100100",
                "남성슈즈": "https://home.29cm.co.kr/best-products?period=HOURLY&ranking=POPULARITY&gender=F&age=30&categoryLargeCode=274100100"
            }

            self.log(f"크롤링 시작: 키워드='{keyword}', 카테고리='{category}', 개수={count}")
            
            async with async_playwright() as p:
                # 브라우저 실행
                browser = await p.chromium.launch(
                    headless=True,
                    args=["--no-sandbox", "--disable-dev-shm-usage"]
                )
                context = await browser.new_context(
                    user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                    viewport={'width': 1280, 'height': 800}
                )
                page = await context.new_page()
                
                target_url = ""
                file_prefix = ""
                
                # URL 결정 로직
                if category and category in category_urls:
                    target_url = category_urls[category]
                    self.log(f"카테고리 베스트 접속: {category}")
                    file_prefix = f"29cm_{category}"
                else:
                    # 키워드 검색
                    if not keyword:
                         keyword = "베스트" 
                    target_url = f"https://www.29cm.co.kr/search/{keyword}"
                    self.log(f"키워드 검색 접속: {keyword}")
                    file_prefix = f"29cm_{keyword}"

                await page.goto(target_url, wait_until='networkidle', timeout=60000)
                await asyncio.sleep(2)
                
                # 스크롤
                await page.mouse.wheel(0, 1000)
                await asyncio.sleep(1)

                # 상품 목록 추출 (개선된 선택자)
                unique_urls = []
                target_items = []
                
                # 재시도 로직
                max_retries = 3
                for attempt in range(max_retries):
                    self.log(f"상품 목록 요소를 찾는 중... (시도 {attempt+1}/{max_retries})")
                    
                    product_elements = []
                    
                    # 1. /product/ 링크 포함 (일반적인 패턴)
                    elements_product = await page.query_selector_all('a[href*="/product/"]')
                    product_elements.extend(elements_product)
                    
                    # 2. /catalog/ 링크 포함 (기존 패턴)
                    elements_catalog = await page.query_selector_all('a[href*="/catalog/"]')
                    product_elements.extend(elements_catalog)
                    
                    # 3. 29cm 특정 클래스 패턴 (ewptmlp5 등 - 동적일 수 있으므로 href 위주로)
                    # 베스트 페이지 구조: div > a (href에 숫자 포함)
                    
                    if not product_elements:
                        # 더 넓은 범위로 검색 (숫자가 포함된 href)
                        all_links = await page.query_selector_all('a[href]')
                        for link in all_links:
                            href = await link.get_attribute('href')
                            if href and ('/product/' in href or '/catalog/' in href):
                                product_elements.append(link)
                    
                    if product_elements:
                        self.log(f"상품 링크 후보 {len(product_elements)}개 발견")
                        
                        for el in product_elements:
                            href = await el.get_attribute('href')
                            if not href: continue
                            
                            # URL 정규화
                            if href.startswith('//'):
                                full_url = f"https:{href}"
                            elif href.startswith('/'):
                                full_url = f"https://www.29cm.co.kr{href}"
                            else:
                                full_url = href
                                
                            # 유효성 검사 (상품 페이지인지)
                            if '/product/' in full_url or '/catalog/' in full_url:
                                # 중복 제거
                                clean_url = full_url.split('?')[0] # 파라미터 제외하고 비교
                                if clean_url not in unique_urls:
                                    unique_urls.append(clean_url)
                                    target_items.append({'url': full_url})
                                    if len(target_items) >= count * 2: # 충분히 수집
                                        break
                                        
                        if len(target_items) > 0:
                            break # 성공
                    
                    # 실패 시 대기 후 재시도
                    await asyncio.sleep(2)
                    # 스크롤 조금 더
                    await page.mouse.wheel(0, 500)
                
                # 요청한 개수만큼 자르기
                target_items = target_items[:count]
                self.log(f"상품 목록 추출 완료: {len(target_items)}개 (목표: {count}개)")
                
                results = []
                import random
                
                # 상세 페이지 순회
                for rank, item in enumerate(target_items, start=1):
                    url = item['url']
                    self.log(f"[{rank}/{len(target_items)}] 상세 정보 수집 중... {url.split('/catalog/')[-1]}")
                    
                    new_page = await context.new_page()
                    try:
                        await new_page.goto(url, wait_until='domcontentloaded', timeout=30000)
                        await asyncio.sleep(random.uniform(1.0, 2.0))
                        
                        # 데이터 수집
                        # 1. 상품명
                        name_elem = await new_page.query_selector('#pdp_product_name')
                        product_name = await name_elem.inner_text() if name_elem else "수집 실패"
                        
                        # 2. 브랜드
                        brand_elem = await new_page.query_selector('a[href*="/brand/"] h3')
                        if not brand_elem:
                             brand_elem = await new_page.query_selector('a[href*="/brand/"][translate="no"]')
                        product_brand = await brand_elem.inner_text() if brand_elem else "수집 실패"
                        
                        # 3. 가격
                        price_elem = await new_page.query_selector('#pdp_product_price')
                        product_price = await price_elem.inner_text() if price_elem else "수집 실패"
                        
                        # 4. 판매자 정보
                        seller_name = ""
                        seller_address = ""
                        contact = ""
                        business_number = ""
                        
                        try:
                            rows = await new_page.query_selector_all('table tr')
                            for row in rows:
                                th_el = await row.query_selector('th')
                                td_el = await row.query_selector('td')
                                if th_el and td_el:
                                    header = (await th_el.inner_text()).replace(" ", "")
                                    value = (await td_el.inner_text()).strip()
                                    
                                    if "상호" in header or "판매자" in header:
                                        if not seller_name: seller_name = value
                                    elif "주소" in header or "소재지" in header:
                                        if not seller_address: seller_address = value
                                    elif "연락처" in header or "전화번호" in header:
                                        if not contact: contact = value
                                    elif "사업자" in header and "번호" in header:
                                        if not business_number: business_number = value
                        except Exception as e:
                            self.log(f"판매자 정보 파싱 오류: {e}")

                        results.append({
                            '순위': rank,
                            '브랜드명': product_brand,
                            '상품명': product_name,
                            '가격': product_price,
                            '판매자 상호': seller_name,
                            '판매자 주소': seller_address,
                            '판매자 주소': seller_address,
                            '연락처': contact,
                            '사업자등록번호': business_number,
                            '상세페이지URL': url
                        })
                        
                    except Exception as e:
                        self.log(f"상품 상세 실패: {e}")
                    finally:
                        await new_page.close()
                    
                    # 진행 상황 업데이트 (GUI가 있다면)
                    # if rank % 5 == 0: self.log(f"진행 중... {rank}개 완료")

                await browser.close()
                
                # 결과 처리
                if results:
                    # 1. Excel 저장
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"{file_prefix}_{timestamp}.xlsx"
                    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                    output_dir = os.path.join(base_dir, "results")
                    if not os.path.exists(output_dir):
                        os.makedirs(output_dir)
                    filepath = os.path.join(output_dir, filename)
                    

                    # 엑셀 저장 전 데이터 포맷팅 (텍스트 강제 지정)
                    for item in results:
                        # 연락처: 앞자리 0 보존을 위해 ' 붙임
                        if item['연락처'] and not item['연락처'].startswith("'"):
                            item['연락처'] = "'" + str(item['연락처'])
                        # 가격: 텍스트로 보존 (선택사항이나 안전을 위해)
                        if item['가격'] and not item['가격'].startswith("'"):
                            item['가격'] = "'" + str(item['가격'])

                    df = pd.DataFrame(results)
                    df.to_excel(filepath, index=False, engine='openpyxl')
                    
                    self.log(f"\n[완료] 파일 저장됨: {filepath}")
                    
                    # 2. TSV 출력 (로그창에 표시하여 복사 가능하게 함)
                    tsv_lines = []
                    headers = ["순위", "브랜드명", "상품명", "가격", "판매자 상호", "판매자 주소", "연락처", "사업자등록번호", "상세페이지URL"]
                    tsv_lines.append("\t".join(headers))
                    
                    for item in results:
                        row = [
                            str(item['순위']),
                            str(item['브랜드명']),
                            str(item['상품명']),
                            str(item['가격']),
                            str(item['판매자 상호']),
                            str(item['판매자 주소']),
                            str(item['연락처']),
                            str(item.get('사업자등록번호', '-')),
                            str(item['상세페이지URL'])
                        ]
                        clean_row = [col.replace('\t', ' ').replace('\n', ' ') for col in row]
                        tsv_lines.append("\t".join(clean_row))
                        
                    final_tsv = "\n".join(tsv_lines)
                    self.log("\n[복사 붙여넣기용 TSV 데이터]")
                    self.log(final_tsv)
                    self.log("=========================================")
                    
                    messagebox.showinfo("완료", f"크롤링이 완료되었습니다.\n{len(results)}개 상품이 저장되었습니다.")
                else:
                    self.log("수집된 상품이 없습니다.")
                
                return results
                    
        except Exception as e:
            error_msg = f"크롤링 오류: {str(e)}"
            self.log(error_msg)
            import traceback
            self.log(traceback.format_exc())
            
        finally:
            self.is_crawling = False
            self.start_button.config(state='normal')
            self.progress_bar.stop()
            self.progress_var.set("완료")


def main():
    """메인 함수"""
    root = tk.Tk()
    app = CrawlerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

