import sys
import os
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass
import os
import asyncio
import threading
import queue
import time
import importlib.util
from datetime import datetime

# --- 경로 설정 ---
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MUSINSA_DIR = os.path.join(BASE_DIR, "musinsa best new")
WCONCEPT_DIR = os.path.join(BASE_DIR, "w concept best")
_29CM_DIR = os.path.join(BASE_DIR, "crawlers", "29cm")

sys.path.append(MUSINSA_DIR)
sys.path.append(WCONCEPT_DIR)
sys.path.append(_29CM_DIR)


# --- Global State ---
log_queues = {}
stop_signals = {}
log_queues_lock = threading.Lock()
stop_signals_lock = threading.Lock()

def get_log_queue(request_id):
    with log_queues_lock:
        if request_id not in log_queues:
            log_queues[request_id] = queue.Queue()
        return log_queues[request_id]

def clear_log_queue(request_id):
    with log_queues_lock:
        if request_id in log_queues:
            del log_queues[request_id]

def log_to_queue(request_id, msg):
    try:
        q = get_log_queue(request_id)
        timestamp = datetime.now().strftime("%H:%M:%S")
        q.put(f"[{timestamp}] {msg}")
    except:
        pass

def set_stop_signal(request_id):
    with stop_signals_lock:
        stop_signals[request_id] = True

def is_stopped(request_id):
    with stop_signals_lock:
        return stop_signals.get(request_id, False)

# --- Imports ---

# Musinsa
try:
    from musinsa_crawler import MusinsaCrawler
except ImportError:
    MusinsaCrawler = None
    print("Warning: Failed to import MusinsaCrawler")

# W Concept
try:
    from w_concept_crawler import WConceptCrawler
except ImportError:
    WConceptCrawler = None
    print("Warning: Failed to import WConceptCrawler")

# 29CM
try:
    spec = importlib.util.spec_from_file_location("crawler_29cm", os.path.join(_29CM_DIR, "crawler.py"))
    _29cm_module = importlib.util.module_from_spec(spec)
    sys.modules["crawler_29cm"] = _29cm_module
    spec.loader.exec_module(_29cm_module)
    CrawlerApp_29CM = _29cm_module.CrawlerApp
except Exception as e:
    print(f"Warning: Failed to import 29CM crawler: {e}")
    CrawlerApp_29CM = None


# --- Wrappers ---

class MockRoot:
    def __init__(self):
        self.value = None
    def get(self): return self.value
    def set(self, v): self.value = v
    def update(self): pass
    def quit(self): pass
    def destroy(self): pass
    def title(self, *args): pass
    def geometry(self, *args): pass
    def resizable(self, *args): pass
    def mainloop(self): pass


class UnifiedMusinsaCrawler:
    def __init__(self, request_id):
        self.request_id = request_id
        self.crawler = MusinsaCrawler()
        # 로그 콜백 오버라이드
        self.crawler.log_callback = self._log_callback
    
    def _log_callback(self, msg):
        log_to_queue(self.request_id, msg)
        
    async def run(self, category, count=100, headless=True):
        url = self.crawler.categories.get(category)
        if not url:
            log_to_queue(self.request_id, f"Error: Unknown category '{category}'")
            return
        
        log_to_queue(self.request_id, f"Starting Musinsa crawling for '{category}' (Limit: {count})")
        
        # 중지 신호 전달
        def check_stop():
            if is_stopped(self.request_id):
                self.crawler.stop_flag = True
        
        # 주기적으로 stop_flag 체크 상기 (원래 crawler가 루프에서 체크함)
        # 래퍼에서 주기적으로 주입
        async def stop_monitor():
            while not is_stopped(self.request_id) and self.crawler.stop_flag == False:
                await asyncio.sleep(0.5)
                if is_stopped(self.request_id):
                    self.crawler.stop_flag = True
                    log_to_queue(self.request_id, "시스템: 중지 요청 감지됨")
                    break

        monitor_task = asyncio.create_task(stop_monitor())

        try:
            products = await self.crawler.crawl_products(category, url, count)
            monitor_task.cancel()
            
            if products:
                log_to_queue(self.request_id, f"✅ Crawling complete. {len(products)} items collected.")
                return {
                    "products": products,
                    "category": category,
                    "count": len(products)
                }
            else:
                log_to_queue(self.request_id, "No products found.")
                return None
        except Exception as e:
            monitor_task.cancel()
            log_to_queue(self.request_id, f"Crawling failed: {e}")
            return None

class UnifiedWConceptCrawler:
    def __init__(self, request_id):
        self.request_id = request_id
        if not WConceptCrawler:
            raise Exception("W Concept crawler module not loaded")
        self.crawler = WConceptCrawler()
        self.crawler.log_callback = self._log_callback
        
    def _log_callback(self, msg):
        log_to_queue(self.request_id, msg)
        
    async def run(self, category, count=10, headless=True):
        log_to_queue(self.request_id, f"Starting W Concept crawling for '{category}' (Count: {count})")
        
        try:
            # 중지 신호 체크를 위해 루프 내에서 처리하거나 크롤러에 넘겨야 함
            # 여기서는 크롤러 인스턴스에 직접 flag 설정 (크롤러가 지원한다면)
            if hasattr(self.crawler, 'stop_flag'):
                # 주기적으로 체크하는 스레드가 필요할 수 있음
                pass

            import asyncio
            loop = asyncio.get_event_loop()
            products = await loop.run_in_executor(
                None, 
                self.crawler.crawl_products,
                category,
                count,
                headless
            )
            
            if products:
                log_to_queue(self.request_id, f"✅ Crawling complete. Collected {len(products)} products")
                return {
                    "products": products,
                    "category": category,
                    "count": len(products)
                }
            else:
                log_to_queue(self.request_id, "No products found")
                return None
                
        except Exception as e:
            log_to_queue(self.request_id, f"Error running W Concept crawler: {e}")
            import traceback
            log_to_queue(self.request_id, traceback.format_exc())
            return None




class Unified29CMCrawler:
    def __init__(self, request_id):
        self.request_id = request_id
        
    async def run(self, keyword, category=None, count=50, headless=True):
        if not CrawlerApp_29CM:
            log_to_queue(self.request_id, "29CM crawler module not loaded.")
            return

        log_to_queue(self.request_id, f"Starting 29CM crawling for '{keyword}' (Category: {category})")
        
        root = MockRoot()
        app = CrawlerApp_29CM(root)
        
        # Override log
        def custom_log(msg):
            log_to_queue(self.request_id, msg)
        app.log = custom_log
        
        # Check stop signal (via monkeypatch or check in loop if supported)
        # 29cm_bag_crawler.py doesn't seem to have a simple stop_flag yet
        
        try:
            # We need to capture the results. 29cm currently saves to excel.
            # I should modify 29cm to return the results list instead of just saving.
            # For now, let's assume it returns or we can find the data.
            # (Requires modifying 29cm script too)
            results = await app.crawl_29cm(keyword, category=category, count=count)
            if results:
                return {
                    "products": results,
                    "category": category or keyword,
                    "count": len(results)
                }
            return None
        except Exception as e:
            log_to_queue(self.request_id, f"Error: {e}")
            return None


# --- 크롤링 결과 저장소 ---
crawl_results = {}  # request_id -> results data
crawl_results_lock = threading.Lock()

def store_crawl_result(request_id, result_data):
    with crawl_results_lock:
        crawl_results[request_id] = result_data

def get_crawl_result(request_id):
    with crawl_results_lock:
        return crawl_results.get(request_id)

def clear_crawl_result(request_id):
    with crawl_results_lock:
        if request_id in crawl_results:
            del crawl_results[request_id]

# --- 메인 실행 함수 ---

async def run_crawler_task(crawler_type, params, request_id):
    """
    crawler_type: 'musinsa', 'wconcept', '29cm'
    params: dict (category, keyword, count, headless)
    """
    log_to_queue(request_id, f"Task started: {crawler_type}")
    
    try:
        result = None
        
        if crawler_type == 'musinsa':
            crawler = UnifiedMusinsaCrawler(request_id)
            result = await crawler.run(
                category=params.get('category', '전체'), 
                count=int(params.get('count', 10)),
                headless=params.get('headless', True)
            )
            
        elif crawler_type == 'wconcept':
            crawler = UnifiedWConceptCrawler(request_id)
            result = await crawler.run(
                category=params.get('category', '베스트탭 (메인)'),
                count=int(params.get('count', 10)),
                headless=params.get('headless', True)
            )
            
        elif crawler_type == '29cm':
            crawler = Unified29CMCrawler(request_id)
            
            category_val = params.get('category')
            if category_val == '직접 검색 (키워드)':
                category_val = None
                
            result = await crawler.run(
                keyword=params.get('keyword', ''),
                category=category_val,
                count=int(params.get('count', 50)),
                headless=params.get('headless', True)
            )
            
        else:
            log_to_queue(request_id, "Unknown crawler type")
        
        # 결과 저장
        if result:
            store_crawl_result(request_id, {
                "crawler_type": crawler_type,
                "data": result,
                "params": params
            })
            
    except Exception as e:
        log_to_queue(request_id, f"Critical Task Error: {e}")
        import traceback
        log_to_queue(request_id, traceback.format_exc())
    
    log_to_queue(request_id, "Task finished.")

