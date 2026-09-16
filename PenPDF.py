import webview
from webview.dom import DOMEventHandler
import os
import fitz  # PyMuPDF 라이브러리 (pip install PyMuPDF)
import base64
import uuid
import json
import tempfile
import winreg
import threading
import shutil
import gc

try:
    import pptx
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

class PdfApi:
    def __init__(self):
        self._window = None
        self.app_version = "ver2026.09.15"
        self.force_close = False  
        
        # [핵심] 암호(unlocked_passwords) 저장소를 완전히 삭제했습니다!
        self.editor_loaded_files = set()
        self.doc_cache = {} 
        self.progress = {"percent": 0, "message": ""}

        self.sys_fonts = {}
        self.korean_fonts = {
            "Malgun Gothic": "맑은 고딕", "Malgun Gothic Bold": "맑은 고딕 (굵게)",
            "Gulim": "굴림", "GulimChe": "굴림체", "Dotum": "돋움", "DotumChe": "돋움체",
            "Batang": "바탕", "BatangChe": "바탕체", "Gungsuh": "궁서", "GungsuhChe": "궁서체",
            "HYGothic-Extra": "HY견고딕", "HYMyeongJo-Extra": "HY견명조",
            "NanumGothic": "나눔고딕", "NanumMyeongjo": "나눔명조", "NanumSquare": "나눔스퀘어"
        }
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as key:
                for i in range(winreg.QueryInfoKey(key)[1]):
                    name, val, _ = winreg.EnumValue(key, i)
                    clean_name = name.split(" (")[0].strip()
                    if not os.path.isabs(val):
                        val = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts', val)
                    self.sys_fonts[clean_name] = val
        except: pass

    # [신규] 자바스크립트가 주기적으로 칠판을 확인하는 함수
    def get_progress(self):
        return getattr(self, 'progress', {"percent": 0, "message": ""})

    def update_progress(self, percent, message):
        self.progress = {"percent": percent, "message": message}
        if self._window:
            # 특수문자 충돌 방지 처리 후 자바스크립트 함수 강제 실행!
            safe_msg = str(message).replace("'", "\\'").replace("\n", "\\n")
            try:
                self._window.evaluate_js(f"if(window.updateProgressUI) window.updateProgressUI({percent}, '{safe_msg}');")
            except: pass

    def clear_loaded_files(self):
        self.editor_loaded_files.clear()
        for doc in self.doc_cache.values():
            try: doc.close()
            except: pass
        self.doc_cache.clear()
        return {"success": True}

    # =================================================================
    # [무(無)저장 보안 엔진] 암호를 기억하지 않고 메모리 복호화만 진행합니다.
    # =================================================================
    def submit_password(self, path, password, pending_files, action_type):
        try:
            norm_path = os.path.normcase(os.path.abspath(path))
            doc = fitz.open(norm_path)
            if doc.authenticate(password):
                # [선생님 요청 반영] 암호 문자열을 절대 저장하지 않습니다.
                # 대신, 메모리에서 암호를 완전히 파괴한 순수한 복사본을 생성해 임시 캐시에 띄웁니다.
                temp_bytes = doc.tobytes(encryption=fitz.PDF_ENCRYPT_NONE)
                doc.close()
                decrypted_doc = fitz.open("pdf", temp_bytes)
                
                if norm_path in self.doc_cache:
                    self.doc_cache[norm_path].close()
                self.doc_cache[norm_path] = decrypted_doc
                
                if action_type == 'add_pdf': return self.add_pdf_for_edit(file_paths=pending_files.get('sources', []), saved_pages=pending_files.get('saved', []))
                elif action_type == 'drop_pdf': return self.add_dropped_pdf_for_edit(file_paths=pending_files.get('sources', []))
                elif action_type == 'load_project': return self.load_project(resume_data=pending_files)
                elif action_type == 'slideshow': return self.load_pdf_for_slideshow(resume_data=pending_files)
                elif action_type == 'slideshow_direct': return self.load_pdf_for_slideshow_direct(file_path=pending_files.get('path'))
                elif action_type in ['compress', 'pptx']: return {"success": True, "path": norm_path, "name": pending_files.get('name'), "size": pending_files.get('size')}
                elif action_type in ['thumbnail_resume', 'highres_resume']: return {"success": True}
            else:
                doc.close()
                return {"success": False, "error": "PASSWORD_INCORRECT"}
        except Exception as e: return {"success": False, "error": str(e)}

    # =================================================================
    # 공통 유틸리티
    # =================================================================
    def get_app_info(self): return {"success": True, "version": self.app_version}
    def set_window(self, window): self._window = window
    def confirm_exit(self): self.force_close = True; self._window.destroy()

    # ▼▼▼ 여기에 아래 4줄을 추가하세요 ▼▼▼
    def open_github(self):
        import webbrowser
        webbrowser.open("https://github.com/seomyeongho/PenPDF")
        return {"success": True}
    # ▲▲▲ 추가 끝 ▲▲▲

    def get_system_fonts(self):
        fonts_list = []
        for eng_name in self.sys_fonts.keys():
            display_name = self.korean_fonts.get(eng_name)
            if display_name: fonts_list.append({"original": eng_name, "display": f"[한글] {display_name}"})
            else:
                if any(x in eng_name for x in ['KR', 'Gothic', 'Myeongjo', 'MyeongJo', 'Batang']): fonts_list.append({"original": eng_name, "display": f"[유력] {eng_name}"})
                else: fonts_list.append({"original": eng_name, "display": eng_name})
        fonts_list.sort(key=lambda x: (0 if "[한글]" in x["display"] else 1, x["display"]))
        if not fonts_list: fonts_list = [{"original": "Malgun Gothic", "display": "[한글] 맑은 고딕"}]
        return {"success": True, "fonts": fonts_list}
    def _format_size(self, size_in_bytes):
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_in_bytes < 1024.0: return f"{size_in_bytes:.2f} {unit}"
            size_in_bytes /= 1024.0
        return f"{size_in_bytes:.2f} TB"

    # =================================================================
    # 1. 파일 불러오기 (보안 철저 확인)
    # =================================================================
    def select_file(self, file_type='pdf', action_tab=''):
        if file_type == 'pdf': file_types = ('PDF 파일 (*.pdf)', '모든 파일 (*.*)')
        else: file_types = ('모든 파일 (*.*)',)
        result = self._window.create_file_dialog(webview.FileDialog.OPEN, allow_multiple=False, file_types=file_types)
        if result and len(result) > 0:
            path = os.path.normcase(os.path.abspath(result[0]))
            name = os.path.basename(path)
            size = self._format_size(os.path.getsize(path))
            
            if file_type == 'pdf' and action_tab not in ['', 'security']:
                if path in self.doc_cache: doc = self.doc_cache[path]
                else: doc = fitz.open(path)
                
                if doc.is_encrypted and not doc.authenticate(""):
                    if path not in self.doc_cache: doc.close()
                    return {"success": False, "error": "PASSWORD_REQUIRED", "path": path, "filename": name, "pending_files": {"path": path, "name": name, "size": size}}
                if path not in self.doc_cache: doc.close()
            return {"success": True, "path": path, "name": name, "size": size}
        return {"success": False, "error": "취소"}

    def get_dropped_file_info(self, file_path, action_tab=''):
        try:
            if not file_path.lower().endswith('.pdf'): return {"success": False, "error": "PDF 파일만 지원합니다."}
            norm_path = os.path.normcase(os.path.abspath(file_path))
            name = os.path.basename(norm_path)
            size = self._format_size(os.path.getsize(norm_path))
            
            if action_tab not in ['', 'security']:
                if norm_path in self.doc_cache: doc = self.doc_cache[norm_path]
                else: doc = fitz.open(norm_path)
                if doc.is_encrypted and not doc.authenticate(""):
                    if norm_path not in self.doc_cache: doc.close()
                    return {"success": False, "error": "PASSWORD_REQUIRED", "path": norm_path, "filename": name, "pending_files": {"path": norm_path, "name": name, "size": size}}
                if norm_path not in self.doc_cache: doc.close()
            return {"success": True, "path": norm_path, "name": name, "size": size}
        except Exception as e: return {"success": False, "error": str(e)}

    def add_pdf_for_edit(self, file_paths=None, saved_pages=None):
        try:
            if file_paths is None:
                result = self._window.create_file_dialog(webview.FileDialog.OPEN, file_types=('PDF 파일 (*.pdf)',), allow_multiple=True)
                if not result: return {"success": False, "error": "취소"}
                file_paths = result if isinstance(result, (list, tuple)) else [result]
            
            if saved_pages is None: saved_pages = []

            for i, file_path in enumerate(file_paths):
                norm_path = os.path.normcase(os.path.abspath(file_path))
                
                if norm_path in self.doc_cache: temp_doc = self.doc_cache[norm_path]
                else: temp_doc = fitz.open(norm_path)
                
                if temp_doc.is_encrypted and not temp_doc.authenticate(""):
                    if norm_path not in self.doc_cache: temp_doc.close()
                    return {
                        "success": False, "error": "PASSWORD_REQUIRED", 
                        "path": norm_path, "filename": os.path.basename(norm_path),
                        "pending_files": {"sources": file_paths[i:], "saved": saved_pages}
                    }
                
                self.editor_loaded_files.add(norm_path)
                if norm_path not in self.doc_cache: self.doc_cache[norm_path] = temp_doc
                    
                doc = self.doc_cache[norm_path]
                for p_idx in range(len(doc)):
                    saved_pages.append({
                        'id': str(uuid.uuid4()), 'source': norm_path, 'filename': os.path.basename(norm_path), 'page_index': p_idx
                    })
            return {"success": True, "pages": saved_pages}
        except Exception as e: return {"success": False, "error": str(e)}

    def add_dropped_pdf_for_edit(self, file_paths):
        try:
            valid_paths = [p for p in file_paths if p.lower().endswith('.pdf')]
            if not valid_paths: return {"success": False, "error": "PDF 파일만 지원합니다."}
            return self.add_pdf_for_edit(file_paths=valid_paths)
        except Exception as e: return {"success": False, "error": f"파일 드롭 오류:\n{str(e)}"}

    # =================================================================
    # 2. 대용량 스트리밍 수신 엔진
    # =================================================================
    def init_stream(self, file_id, filename):
        if not hasattr(self, 'upload_streams'): self.upload_streams = {}
        temp_dir = os.path.join(tempfile.gettempdir(), "PenPDF_Drops")
        os.makedirs(temp_dir, exist_ok=True)
        temp_path = os.path.join(temp_dir, filename)
        self.upload_streams[file_id] = open(temp_path, 'wb')
        return {"success": True, "path": temp_path}
    def append_stream(self, file_id, b64_chunk):
        if hasattr(self, 'upload_streams') and file_id in self.upload_streams:
            self.upload_streams[file_id].write(base64.b64decode(b64_chunk))
        return True
    def close_stream(self, file_id):
        if hasattr(self, 'upload_streams') and file_id in self.upload_streams:
            self.upload_streams[file_id].close()
            del self.upload_streams[file_id]
        return True
    def process_streamed_files(self, paths, target_tab):
        try:
            if target_tab == 'editor': return self.add_dropped_pdf_for_edit(paths)
            elif target_tab == 'slideshow': return self.load_pdf_for_slideshow_direct(paths[0])
            elif target_tab in ['compress', 'pptx', 'security']: return self.get_dropped_file_info(paths[0], target_tab)
            return {"success": False, "error": "알 수 없는 탭입니다."}
        except Exception as e: return {"success": False, "error": str(e)}

    # =================================================================
    # 3. 프로젝트 파일 열기 / 저장
    # =================================================================
    def load_project(self, resume_data=None):
        try:
            if resume_data is None:
                result = self._window.create_file_dialog(webview.FileDialog.OPEN, file_types=('PenPDF Project (*.ppdf)', '모든 파일 (*.*)'))
                if not result: return {"success": False, "error": "취소"}
                with open(result[0], 'r', encoding='utf-8') as f: project_data = json.load(f)
                grid_data = project_data.get('grid', [])
                for page in grid_data:
                    if 'thumbnail' in page: del page['thumbnail']
                sources = list(set(p['source'] for p in grid_data))
                file_path = result[0]
            else:
                project_data = resume_data['project_data']
                grid_data = project_data.get('grid', [])
                sources = resume_data['pending_sources']
                file_path = resume_data['path']

            for i, src in enumerate(sources):
                norm_src = os.path.normcase(os.path.abspath(src))
                if os.path.exists(norm_src):
                    if norm_src in self.doc_cache: temp_doc = self.doc_cache[norm_src]
                    else: temp_doc = fitz.open(norm_src)
                    
                    if temp_doc.is_encrypted and not temp_doc.authenticate(""):
                        if norm_src not in self.doc_cache: temp_doc.close()
                        return {"success": False, "error": "PASSWORD_REQUIRED", "path": norm_src, "filename": os.path.basename(norm_src), "pending_files": {"project_data": project_data, "pending_sources": sources[i:], "path": file_path}}
                    if norm_src not in self.doc_cache: temp_doc.close()

            light_project_data = {"grid": grid_data, "edits": project_data.get('edits', {}), "globals": project_data.get('globals', {})}
            return {"success": True, "data": json.dumps(light_project_data), "path": file_path}
        except Exception as e: return {"success": False, "error": str(e)}

    def save_project(self, project_data_str, current_path=None, force_dialog=False):
        try:
            if not current_path or force_dialog:
                save_result = self._window.create_file_dialog(webview.FileDialog.SAVE, save_filename="MyProject.ppdf", file_types=('PenPDF Project (*.ppdf)', '모든 파일 (*.*)'))
                if not save_result: return {"success": False, "error": "취소"}
                current_path = save_result[0]
            with open(current_path, 'w', encoding='utf-8') as f: f.write(project_data_str)
            return {"success": True, "saved_path": current_path}
        except Exception as e: return {"success": False, "error": str(e)}

    # =================================================================
    # 4. 이미지 렌더링 (썸네일 및 편집용)
    # =================================================================
    def get_page_thumbnail(self, source, page_index):
        try:
            norm_path = os.path.normcase(os.path.abspath(source))
            if norm_path not in self.doc_cache: self.doc_cache[norm_path] = fitz.open(norm_path)
            doc = self.doc_cache[norm_path]
            
            if doc.is_encrypted and not doc.authenticate(""):
                return {"success": False, "error": "PASSWORD_REQUIRED", "path": norm_path, "filename": os.path.basename(norm_path)}
                
            page = doc.load_page(page_index)
            pix = page.get_pixmap(matrix=fitz.Matrix(0.3, 0.3))
            b64 = base64.b64encode(pix.tobytes("png")).decode('utf-8')
            return {"success": True, "image": f"data:image/png;base64,{b64}"}
        except Exception as e: return {"success": False, "error": str(e)}

    def get_highres_image(self, source, page_index, crop_box=None):
        if not source: return {"success": False, "error": "파일을 찾을 수 없습니다."}
        try:
            src = os.path.normcase(os.path.abspath(source.replace('file:///', '').replace('file://', '')))
            if src in self.doc_cache:
                doc = fitz.open("pdf", self.doc_cache[src].tobytes())
            else:
                doc = fitz.open(src)
                if doc.is_encrypted:
                    if not doc.authenticate(""):
                        doc.close()
                        return {"success": False, "error": "보안 해제가 필요합니다."}

            page = doc[page_index]
            
            clip_rect = None
            if crop_box:
                w = page.rect.width
                h = page.rect.height
                clip_rect = fitz.Rect(
                    page.rect.x0 + crop_box['left'] * w,
                    page.rect.y0 + crop_box['top'] * h,
                    page.rect.x0 + crop_box['right'] * w,
                    page.rect.y0 + crop_box['bottom'] * h
                )
                
            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False, clip=clip_rect)
            
            # [화질 복구 완료] 마음대로 jpeg로 낮췄던 것을 다시 고화질 무손실 png로 원상 복구했습니다!
            img_bytes = pix.tobytes("png")
            b64 = base64.b64encode(img_bytes).decode('utf-8')
            doc.close()
            return {"success": True, "image": "data:image/png;base64," + b64}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # =================================================================
    # 5. 슬라이드쇼 화면 로딩
    # =================================================================
    def load_pdf_for_slideshow(self, resume_data=None):
        try:
            if resume_data is None:
                result = self._window.create_file_dialog(webview.FileDialog.OPEN, file_types=('PDF 파일 (*.pdf)', '모든 파일 (*.*)'))
                if not result: return {"success": False, "error": "취소"}
                file_path = result[0]
            else: file_path = resume_data['path']
            
            norm_path = os.path.normcase(os.path.abspath(file_path))
            if norm_path in self.doc_cache: doc = self.doc_cache[norm_path]
            else: doc = fitz.open(norm_path)
            
            if doc.is_encrypted and not doc.authenticate(""):
                if norm_path not in self.doc_cache: doc.close()
                return {"success": False, "error": "PASSWORD_REQUIRED", "path": norm_path, "filename": os.path.basename(norm_path), "pending_files": {"path": norm_path}}
            
            pages = []
            for i in range(len(doc)):
                page = doc.load_page(i)
                pix = page.get_pixmap(matrix=fitz.Matrix(0.2, 0.2), alpha=False)
                img_base64 = base64.b64encode(pix.tobytes("jpeg")).decode('utf-8')
                pages.append({"source": norm_path, "page_index": i, "thumbnail": f"data:image/jpeg;base64,{img_base64}"})
            
            name = os.path.basename(norm_path)
            size = self._format_size(os.path.getsize(norm_path))
            if norm_path not in self.doc_cache: doc.close()
            return {"success": True, "pages": pages, "name": name, "size": size, "path": norm_path}
        except Exception as e: return {"success": False, "error": str(e)}

    def load_pdf_for_slideshow_direct(self, file_path):
        try:
            if not file_path.lower().endswith('.pdf'): return {"success": False, "error": "PDF 파일만 지원합니다."}
            norm_path = os.path.normcase(os.path.abspath(file_path))
            if norm_path in self.doc_cache: doc = self.doc_cache[norm_path]
            else: doc = fitz.open(norm_path)
            
            if doc.is_encrypted and not doc.authenticate(""):
                if norm_path not in self.doc_cache: doc.close()
                return {"success": False, "error": "PASSWORD_REQUIRED", "path": norm_path, "filename": os.path.basename(norm_path), "pending_files": {"path": norm_path}}
            
            pages = []
            for i in range(len(doc)):
                page = doc.load_page(i)
                pix = page.get_pixmap(matrix=fitz.Matrix(0.2, 0.2), alpha=False)
                img_base64 = base64.b64encode(pix.tobytes("jpeg")).decode('utf-8')
                pages.append({"source": norm_path, "page_index": i, "thumbnail": f"data:image/jpeg;base64,{img_base64}"})
            
            name = os.path.basename(norm_path)
            size = self._format_size(os.path.getsize(norm_path))
            if norm_path not in self.doc_cache: doc.close()
            return {"success": True, "pages": pages, "name": name, "size": size, "path": norm_path}
        except Exception as e: return {"success": False, "error": str(e)}

    # [신규] 파일이 선택되면 투명 이미지가 몇 개인지 빠르게 훑어보는 함수
    def check_transparent_images(self, file_path):
        try:
            norm_path = os.path.normcase(os.path.abspath(file_path))
            if norm_path in self.doc_cache:
                doc = fitz.open("pdf", self.doc_cache[norm_path].tobytes())
            else:
                doc = fitz.open(norm_path)
                if doc.is_encrypted:
                    if not doc.authenticate(""):
                        doc.close()
                        return {"success": False, "error": "PASSWORD_REQUIRED", "path": norm_path, "filename": os.path.basename(norm_path)}
                    doc = fitz.open("pdf", doc.tobytes(encryption=fitz.PDF_ENCRYPT_NONE))

            # 압축 함수와 완벽히 동일한 기준으로 마스크 중복 카운트를 방지합니다.
            all_mask_xrefs = set()
            for page_idx in range(len(doc)):
                page = doc.load_page(page_idx)
                for img in page.get_images(full=True):
                    smask_xref = img[1]
                    if smask_xref > 0:
                        all_mask_xrefs.add(smask_xref)

            transparent_images = []
            processed_xrefs = set()
            
            for page_idx in range(len(doc)):
                page = doc.load_page(page_idx)
                for img in page.get_images(full=True):
                    xref = img[0]
                    smask_xref = img[1]
                    
                    if xref in processed_xrefs or xref in all_mask_xrefs: 
                        continue
                    processed_xrefs.add(xref)
                    
                    is_transparent = False
                    if smask_xref > 0: is_transparent = True
                    else:
                        is_mask = doc.xref_get_key(xref, "ImageMask")
                        if is_mask[0] != "null" and is_mask[1] == "true": is_transparent = True
                        else:
                            has_smask = doc.xref_get_key(xref, "SMask")
                            if has_smask[0] != "null": is_transparent = True
                            else:
                                has_mask = doc.xref_get_key(xref, "Mask")
                                if has_mask[0] != "null": is_transparent = True
                                else:
                                    try:
                                        pix = fitz.Pixmap(doc, xref)
                                        if pix.alpha: is_transparent = True
                                        pix = None
                                    except: pass
                    
                    if is_transparent:
                        width, height = img[2], img[3]
                        try:
                            raw_stream = doc.xref_stream_raw(xref)
                            size = len(raw_stream) if raw_stream else (width * height * 3)
                        except AttributeError:
                            stream = doc.xref_stream(xref)
                            size = len(stream) if stream else (width * height * 3)
                            
                        # 상세 정보를 리스트에 차곡차곡 저장합니다.
                        transparent_images.append({
                            "page": page_idx + 1,
                            "width": width,
                            "height": height,
                            "size": size
                        })
            
            doc.close()
            return {"success": True, "count": len(transparent_images), "images": transparent_images}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # =================================================================
    # 6. PDF 압축 / PPTX 변환 / 에디터 저장
    # =================================================================
    # [핵심] 파라미터에 투명 보호 여부와 색상 코드가 추가되었습니다.
    def compress_pdf(self, input_path, level, image_quality, open_folder, protect_transparent=False, transparent_fill_color='#FFFFFF'):
        import math
        if not input_path or not os.path.exists(input_path): return {"success": False, "error": "파일을 찾을 수 없습니다."}
        
        try:
            import io
            from PIL import Image
            has_pil = True
        except ImportError:
            has_pil = False

        try:
            default_save_path = input_path.replace('.pdf', '_compressed.pdf')
            save_result = self._window.create_file_dialog(webview.FileDialog.SAVE, save_filename=os.path.basename(default_save_path), file_types=('PDF 파일 (*.pdf)',))
            if not save_result: return {"success": False, "error": "취소"}
            output_path = save_result[0]
            
            if self._window:
                self._window.evaluate_js("showLoading('PDF 압축 중입니다...\\n(파일 용량에 따라 다소 시간이 소요될 수 있습니다)'); startProgressPolling();")
            self.update_progress(5, "1단계: 문서 구조 파악 및 압축 준비 중...")

            try:
                norm_out = os.path.normcase(os.path.abspath(output_path))
                norm_in = os.path.normcase(os.path.abspath(input_path))
                if norm_out == norm_in or (os.path.exists(norm_out) and os.path.exists(norm_in) and os.path.samefile(norm_out, norm_in)):
                    if self._window: self._window.evaluate_js("stopProgressPolling(); hideLoading();")
                    return {"success": False, "error": "원본 파일 보호 🛡️\n\n원본과 동일한 이름으로 덮어쓸 수 없습니다."}
            except: pass

            original_size = os.path.getsize(input_path)
            norm_in = os.path.normcase(os.path.abspath(input_path))
            
            self.update_progress(20, "2단계: 보안 스캔 및 파일 메모리 로드 중...")
            if norm_in in self.doc_cache:
                doc = fitz.open("pdf", self.doc_cache[norm_in].tobytes())
            else:
                doc = fitz.open(norm_in)
                if doc.is_encrypted:
                    if not doc.authenticate(""):
                        doc.close()
                        if self._window: self._window.evaluate_js("stopProgressPolling(); hideLoading();")
                        return {"success": False, "error": "보안 해제가 필요합니다. 다시 시도해주세요."}
                    temp_bytes = doc.tobytes(encryption=fitz.PDF_ENCRYPT_NONE)
                    doc.close()
                    doc = fitz.open("pdf", temp_bytes)
            
            compression_rate = int(image_quality)
            quality = 100 - compression_rate                    
            
            self.update_progress(50, "3단계: 지능형 무손실 이미지 다이어트 진행 중 (시간이 다소 소요됩니다)...")
            
            details_log = []
            gc.collect()
            
            if quality < 100:
                all_mask_xrefs = set()
                for page_idx in range(len(doc)):
                    page = doc.load_page(page_idx)
                    for img in page.get_images(full=True):
                        smask_xref = img[1]
                        if smask_xref > 0:
                            all_mask_xrefs.add(smask_xref)

                processed_xrefs = set()
                for page_idx in range(len(doc)):
                    page = doc.load_page(page_idx)
                    for img in page.get_images(full=True):
                        xref = img[0]
                        smask_xref = img[1]
                        
                        if xref in processed_xrefs or xref in all_mask_xrefs: 
                            continue
                        processed_xrefs.add(xref)
                        
                        width, height = img[2], img[3]
                        if width == 0 or height == 0: continue
                        
                        try:
                            raw_stream = doc.xref_stream_raw(xref)
                            old_len = len(raw_stream) if raw_stream else (width * height * 3)
                        except AttributeError:
                            stream = doc.xref_stream(xref)
                            old_len = len(stream) if stream else (width * height * 3)
                            
                        log_entry = {
                            "page": page_idx + 1, "width": width, "height": height,
                            "old_size": old_len, "new_size": old_len, "status": "", "ratio": 0
                        }

                        if old_len < 2048:
                            log_entry["status"] = "스킵 (초소형 데이터)"
                            details_log.append(log_entry)
                            continue

                        # 투명도 존재 여부 판별
                        is_transparent = False
                        if smask_xref > 0: is_transparent = True
                        
                        is_mask = doc.xref_get_key(xref, "ImageMask")
                        if is_mask[0] != "null" and is_mask[1] == "true": is_transparent = True
                        
                        has_smask = doc.xref_get_key(xref, "SMask")
                        if has_smask[0] != "null": is_transparent = True

                        has_mask = doc.xref_get_key(xref, "Mask")
                        if has_mask[0] != "null": is_transparent = True

                        try:
                            pix = fitz.Pixmap(doc, xref)
                            if pix.alpha: is_transparent = True
                            
                            # 선생님 옵션에 따라 스킵할지 결정
                            if is_transparent and protect_transparent:
                                log_entry["status"] = "스킵 (투명 배경 보호 - 원본 유지)"
                                details_log.append(log_entry)
                                pix = None
                                continue
                                
                            if pix.n > 3 and not pix.alpha:
                                temp = fitz.Pixmap(fitz.csRGB, pix)
                                pix = temp
                                
                            current_quality = quality
                            new_bytes = None
                            final_w = width
                            final_h = height
                            use_pil = False
                            
                            # [핵심 로직] 보조 엔진(PIL)을 이용해 투명 배경에 색상을 들이붓고 합성합니다!
                            if has_pil:
                                try:
                                    img_mode = "RGBA" if pix.alpha else ("RGB" if pix.n >= 3 else ("L" if pix.n == 1 else "CMYK"))
                                    img_obj = Image.frombytes(img_mode, [pix.width, pix.height], pix.samples)
                                    if img_obj.mode == 'CMYK':
                                        img_obj = img_obj.convert('RGB')
                                    
                                    # 숨겨진 투명도 스텐실(SMask)을 불러와서 이미지에 씌워줍니다.
                                    if smask_xref > 0:
                                        smask_pix = fitz.Pixmap(doc, smask_xref)
                                        smask_img = Image.frombytes("L" if smask_pix.n == 1 else "RGB", [smask_pix.width, smask_pix.height], smask_pix.samples)
                                        if smask_img.mode != "L": smask_img = smask_img.convert("L")
                                        if smask_img.size != img_obj.size: smask_img = smask_img.resize(img_obj.size)
                                        
                                        if img_obj.mode != "RGBA": img_obj = img_obj.convert("RGBA")
                                        img_obj.putalpha(smask_img)
                                    
                                    # 사용자가 지정한 색상으로 도화지를 만들고 그 위에 투명 이미지를 얹습니다.
                                    if is_transparent or img_obj.mode in ('RGBA', 'LA') or (img_obj.mode == 'P' and 'transparency' in img_obj.info):
                                        if img_obj.mode != 'RGBA': img_obj = img_obj.convert('RGBA')
                                        bg_color = transparent_fill_color
                                        background = Image.new('RGB', img_obj.size, bg_color)
                                        background.paste(img_obj, mask=img_obj.split()[3])
                                        img_obj = background # 완벽하게 색상이 채워진 이미지 완성
                                        
                                    scale = math.sqrt(current_quality / 100.0)
                                    if scale < 1.0:
                                        final_w = max(1, int(pix.width * scale))
                                        final_h = max(1, int(pix.height * scale))
                                        try: resample_method = Image.Resampling.LANCZOS
                                        except AttributeError: resample_method = Image.LANCZOS
                                        img_obj = img_obj.resize((final_w, final_h), resample_method)
                                        
                                    byte_arr = io.BytesIO()
                                    img_obj.save(byte_arr, format='JPEG', quality=current_quality, optimize=True)
                                    new_bytes = byte_arr.getvalue()
                                    use_pil = True
                                except Exception:
                                    new_bytes = None
                            
                            # (만약 PIL이 없을 때를 대비한 기본 엔진 백업 로직)
                            if not use_pil:
                                final_w = pix.width
                                final_h = pix.height
                                
                                if is_transparent:
                                    hex_c = transparent_fill_color.lstrip('#')
                                    r, g, b = int(hex_c[0:2], 16), int(hex_c[2:4], 16), int(hex_c[4:6], 16)
                                    bg = fitz.Pixmap(fitz.csRGB, pix.width, pix.height)
                                    bg.set_origin(pix.x, pix.y)
                                    bg.clear_with(r, g, b)
                                    bg.copy(pix, pix.irect)
                                    pix = bg
                                    
                                try: new_bytes = pix.tobytes("jpeg", jpg_quality=current_quality)
                                except TypeError:
                                    try: new_bytes = pix.tobytes("jpeg", quality=current_quality)
                                    except TypeError: new_bytes = pix.tobytes("jpeg")
                            
                            new_len = len(new_bytes)
                            target_len = old_len * (quality / 100.0)
                            retry_quality = current_quality
                            
                            while new_len > target_len and retry_quality > 15:
                                retry_quality -= 5
                                
                                if use_pil:
                                    byte_arr = io.BytesIO()
                                    img_obj.save(byte_arr, format='JPEG', quality=retry_quality, optimize=True)
                                    new_bytes = byte_arr.getvalue()
                                else:
                                    if retry_quality <= 60 and final_w == width and width > 1000:
                                        pix.shrink(1)
                                        final_w = pix.width
                                        final_h = pix.height
                                        
                                    try: new_bytes = pix.tobytes("jpeg", jpg_quality=retry_quality)
                                    except TypeError:
                                        try: new_bytes = pix.tobytes("jpeg", quality=retry_quality)
                                        except TypeError: new_bytes = pix.tobytes("jpeg")
                                            
                                new_len = len(new_bytes)
                                log_entry["status"] = f"상대적 압축({retry_quality}%) & "
                            
                            if new_len >= old_len:
                                log_entry["status"] = "스킵 (최대 압축 도달, 이득 없음)"
                                log_entry["new_size"] = old_len
                                details_log.append(log_entry)
                                pix = None
                                continue

                            is_compressed = False
                            try:
                                doc.update_stream(xref, new_bytes, compress=False)
                            except TypeError:
                                doc.update_stream(xref, new_bytes)
                                is_compressed = True
                                
                            doc.xref_set_key(xref, "Width", str(final_w))
                            doc.xref_set_key(xref, "Height", str(final_h))
                            doc.xref_set_key(xref, "BitsPerComponent", "8")
                            
                            color_space = "/DeviceRGB" if pix.n >= 3 else "/DeviceGray"
                            doc.xref_set_key(xref, "ColorSpace", color_space)
                            
                            if is_compressed:
                                doc.xref_set_key(xref, "Filter", "[/FlateDecode /DCTDecode]")
                            else:
                                doc.xref_set_key(xref, "Filter", "/DCTDecode")
                                
                            doc.xref_set_key(xref, "Decode", "null")
                            doc.xref_set_key(xref, "DecodeParms", "null")
                            
                            # [핵심] 투명도 정보를 제거했으므로, 찌꺼기 메타데이터도 날려버립니다.
                            doc.xref_set_key(xref, "ImageMask", "null")
                            doc.xref_set_key(xref, "Mask", "null")
                            doc.xref_set_key(xref, "SMask", "null")
                            
                            if "상대적 압축" not in log_entry["status"]:
                                log_entry["status"] = f"해상도/화질({current_quality}%) 압축 & 성공"
                            else:
                                log_entry["status"] += "성공"
                                
                            log_entry["new_size"] = new_len
                            log_entry["ratio"] = round((old_len - new_len) / old_len * 100, 1)
                            details_log.append(log_entry)
                                
                            pix = None
                        except Exception as e:
                            log_entry["status"] = f"에러 스킵"
                            details_log.append(log_entry)
                            
            gc.collect()

            self.update_progress(85, "4단계: 불필요한 메타데이터 제거 및 디스크 저장 중...")
            garbage_level = 1 if level == "low" else (3 if level == "medium" else 4)
            
            if level == "high":
                doc.save(output_path, garbage=garbage_level, deflate=True, deflate_images=True, deflate_fonts=True, use_objstms=True, clean=True)
            else:
                doc.save(output_path, garbage=garbage_level, deflate=True, deflate_images=True, deflate_fonts=True, clean=True)
            
            doc.close()
            gc.collect()

            self.update_progress(100, "압축이 완료되었습니다!")

            compressed_size = os.path.getsize(output_path)
            
            # [수정] 압축 후 용량이 같거나 오히려 늘어났다면?
            if compressed_size >= original_size:
                # 헛수고한 결과물은 즉시 삭제해버립니다.
                try:
                    os.remove(output_path)
                except Exception:
                    pass
                
                return {
                    "success": True, 
                    "already_optimized": True, # [신규] UI에서 이 신호를 보고 화면을 다르게 그립니다.
                    "original_size": self._format_size(original_size), 
                    "compressed_size": self._format_size(original_size), 
                    "ratio": "0.0% (이미 최적화됨)", 
                    "saved_path": "",
                    "details": [] 
                }
            else:
                # 정상적으로 압축이 성공했을 때만 파일을 남깁니다.
                ratio = ((original_size - compressed_size) / original_size) * 100 if original_size > 0 else 0
                if open_folder:
                    try: os.startfile(os.path.dirname(output_path))
                    except: pass

                return {
                    "success": True, 
                    "already_optimized": False,
                    "original_size": self._format_size(original_size), 
                    "compressed_size": self._format_size(compressed_size), 
                    "ratio": f"{ratio:.1f}%", 
                    "saved_path": output_path,
                    "details": details_log
                }
                
        except Exception as e: 
            gc.collect()
            return {"success": False, "error": f"압축 오류:\n{str(e)}"}

    def convert_to_pptx(self, input_path, convert_mode, open_folder):
        if not PPTX_AVAILABLE: return {"success": False, "error": "이 기능을 사용하려면 'python-pptx' 모듈이 필요합니다.\npip install python-pptx"}
        if not input_path or not os.path.exists(input_path): return {"success": False, "error": "원본 파일을 찾을 수 없습니다."}
        try:
            default_save_path = input_path.replace('.pdf', '_converted.pptx')
            save_result = self._window.create_file_dialog(webview.FileDialog.SAVE, save_filename=os.path.basename(default_save_path), file_types=('파워포인트 파일 (*.pptx)',))
            if not save_result: return {"success": False, "error": "취소"}
            output_path = save_result[0]

            # [핵심] 사용자가 다이얼로그에서 저장을 누른 직후에 로딩창 띄우기
            if self._window:
                self._window.evaluate_js("showLoading('PPTX로 변환 중입니다...\\n(페이지 수에 따라 1~3분 정도 소요될 수 있습니다)'); startProgressPolling();")
            self.update_progress(5, "1단계: 변환 모듈 및 환경 구성 중...")

            self.update_progress(10, "2단계: 보안 해제 및 메모리 조립 중...")
            norm_in = os.path.normcase(os.path.abspath(input_path))
            if norm_in in self.doc_cache:
                doc = fitz.open("pdf", self.doc_cache[norm_in].tobytes())
            else:
                doc = fitz.open(norm_in)
                if doc.is_encrypted:
                    if not doc.authenticate(""):
                        doc.close()
                        if self._window: self._window.evaluate_js("stopProgressPolling(); hideLoading();")
                        return {"success": False, "error": "보안 해제가 필요합니다. 다시 시도해주세요."}
                    temp_bytes = doc.tobytes(encryption=fitz.PDF_ENCRYPT_NONE)
                    doc.close()
                    doc = fitz.open("pdf", temp_bytes)
                
            prs = pptx.Presentation()
            if len(doc) > 0: prs.slide_width = Pt(doc[0].rect.width); prs.slide_height = Pt(doc[0].rect.height)
            blank_slide_layout = prs.slide_layouts[6] 

            total_pages = len(doc)
            for page_index in range(total_pages):
                current_pct = 15 + int((page_index / total_pages) * 75)
                self.update_progress(current_pct, f"3단계: 페이지 변환 중 ({page_index + 1} / {total_pages} 장)")

                page = doc.load_page(page_index)
                slide = prs.slides.add_slide(blank_slide_layout)
                
                if convert_mode == "image":
                    # [버그 수정 완결] PyMuPDF의 alpha 병합 충돌 에러 원천 차단!
                    # 투명도를 그대로 유지하는 PNG 파일로 저장하여 PPTX 슬라이드 위에 올립니다.
                    # 파워포인트 슬라이드 배경이 흰색이므로 검게 타는 현상도 자연스럽게 완벽히 해결됩니다.
                    pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=True)
                    temp_img_path = os.path.join(tempfile.gettempdir(), f"temp_pptx_{uuid.uuid4().hex}.png")
                    pix.save(temp_img_path)
                    slide.shapes.add_picture(temp_img_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
                    os.remove(temp_img_path)
                    pix = None
                else:
                    text_dict = page.get_text("dict")
                    for block in text_dict.get("blocks", []):
                        if block["type"] == 0:  
                            x0, y0, x1, y1 = block["bbox"]
                            txBox = slide.shapes.add_textbox(Pt(x0), Pt(y0), Pt(max(x1 - x0, 20)), Pt(max(y1 - y0, 15)))
                            tf = txBox.text_frame
                            tf.clear()
                            tf.word_wrap = False 
                            for line_idx, line in enumerate(block.get("lines", [])):
                                p = tf.paragraphs[0] if line_idx == 0 else tf.add_paragraph()
                                for span in line.get("spans", []):
                                    text = span.get("text", "") 
                                    if not text: continue
                                    run = p.add_run()
                                    run.text = text
                                    run.font.size = Pt(span["size"])
                                    color_int = span.get("color", 0)
                                    run.font.color.rgb = RGBColor((color_int >> 16) & 255, (color_int >> 8) & 255, color_int & 255)
                        elif block["type"] == 1: 
                            img_bytes = block.get("image")
                            if img_bytes:
                                try:
                                    # [버그 수정 2] PPTX 파일 손상 원인 차단! 
                                    # 원본 PDF의 기괴한 포맷(JBIG2, JPX 등)을 파워포인트가 좋아하는 완벽한 PNG로 재렌더링
                                    img_pix = fitz.Pixmap(img_bytes)
                                    if img_pix.n - img_pix.alpha > 3: # CMYK인 경우 파워포인트 호환을 위해 RGB로 변환
                                        temp_pix = fitz.Pixmap(fitz.csRGB, img_pix)
                                        img_pix = temp_pix
                                        
                                    temp_img_path = os.path.join(tempfile.gettempdir(), f"temp_img_{uuid.uuid4().hex}.png")
                                    img_pix.save(temp_img_path)
                                    
                                    slide.shapes.add_picture(temp_img_path, Pt(block["bbox"][0]), Pt(block["bbox"][1]), width=Pt(block["bbox"][2]-block["bbox"][0]), height=Pt(block["bbox"][3]-block["bbox"][1]))
                                    os.remove(temp_img_path)
                                    img_pix = None
                                except Exception:
                                    pass # 해독 불가한 찌꺼기 이미지는 과감히 스킵하여 PPTX 전체가 박살나는 것을 막음
            
            self.update_progress(90, "4단계: 변환된 파워포인트 문서 디스크에 저장 중...")
            doc.close()
            prs.save(output_path)
            
            self.update_progress(100, "PPTX 변환 완료!")
            if open_folder:
                try: os.startfile(os.path.dirname(output_path))
                except: pass
            return {"success": True, "saved_path": output_path}
        except Exception as e: return {"success": False, "error": f"PPTX 변환 오류:\n{str(e)}"}

    def _draw_shapes_on_page(self, page, shapes):
        # [수정] 1차 패스: '지우개(Whiteout)' 객체만 쏙쏙 골라서 먼저 파기(Redact) 적용
        has_redaction = False
        
        for s in shapes:
            if s.get('isWhiteout'):
                cw, ch = float(s.get('canvasWidth', 800)), float(s.get('canvasHeight', 1131))
                rx, ry = page.rect.width / max(cw, 1), page.rect.height / max(ch, 1)
                
                if 'centerX' in s and 'centerY' in s: 
                    center_x, center_y = float(s['centerX']) * rx, float(s['centerY']) * ry
                else: 
                    center_x, center_y = (float(s.get('left', 0)) + float(s.get('width', 0)) * abs(float(s.get('scaleX', 1))) / 2.0) * rx, (float(s.get('top', 0)) + float(s.get('height', 0)) * abs(float(s.get('scaleY', 1))) / 2.0) * ry
                
                center = fitz.Point(center_x, center_y)
                hw = float(s.get('width', 0)) * abs(float(s.get('scaleX', 1))) * rx / 2.0
                hh = float(s.get('height', 0)) * abs(float(s.get('scaleY', 1))) * ry / 2.0
                
                fx = -1 if (float(s.get('scaleX', 1)) < 0) ^ bool(s.get('flipX', False)) else 1
                fy = -1 if (float(s.get('scaleY', 1)) < 0) ^ bool(s.get('flipY', False)) else 1
                mat_raw = fitz.Matrix(fx, 0, 0, fy, 0, 0) * fitz.Matrix(float(s.get('angle', 0)))
                mat_transform = fitz.Matrix(1, 0, 0, 1, -center.x, -center.y) * mat_raw * fitz.Matrix(1, 0, 0, 1, center.x, center.y)
                
                rp = [p * mat_transform for p in [fitz.Point(center_x - hw, center_y - hh), fitz.Point(center_x + hw, center_y - hh), fitz.Point(center_x - hw, center_y + hh), fitz.Point(center_x + hw, center_y + hh)]]
                quad = fitz.Quad(rp[0], rp[1], rp[2], rp[3])
                
                # 해당 영역을 영구 삭제 대상(Annotation)으로 지정하고 하얀색(1,1,1)으로 채웁니다.
                page.add_redact_annot(quad, fill=(1,1,1))
                has_redaction = True

        # 지우개 폭탄을 먼저 터뜨려서 원본 데이터를 물리적으로 깔끔하게 지워버립니다.
        if has_redaction:
            page.apply_redactions()

        # [수정] 2차 패스: 안전해진 도화지 위에 사용자가 추가한 일반 도형, 텍스트, 이미지를 올립니다.
        for s in shapes:
            if s.get('isWhiteout'): 
                continue
                
            stype = s.get('type')
            if not stype: continue
            cw, ch = float(s.get('canvasWidth', 800)), float(s.get('canvasHeight', 1131))
            rx, ry = page.rect.width / max(cw, 1), page.rect.height / max(ch, 1)
            
            def hex_to_rgb(hex_str):
                if not hex_str or hex_str == 'transparent' or hex_str == 'rgba(0,0,0,0)': return None
                h = hex_str.lstrip('#')
                return tuple(int(h[i:i+2], 16) / 255 for i in (0, 2, 4)) if len(h) == 6 else (0,0,0)

            f_color = hex_to_rgb(s.get('fill'))
            s_color = hex_to_rgb(s.get('stroke'))
            s_width = float(s.get('strokeWidth', 0)) * rx
            if s_width <= 0: s_color, s_width = None, 0.1
            opacity = float(s.get('opacity', 1.0))
            
            if 'centerX' in s and 'centerY' in s: center_x, center_y = float(s['centerX']) * rx, float(s['centerY']) * ry
            else: center_x, center_y = (float(s.get('left', 0)) + float(s.get('width', 0)) * abs(float(s.get('scaleX', 1))) / 2.0) * rx, (float(s.get('top', 0)) + float(s.get('height', 0)) * abs(float(s.get('scaleY', 1))) / 2.0) * ry
            
            center = fitz.Point(center_x, center_y)
            hw = float(s.get('width', 0)) * abs(float(s.get('scaleX', 1))) * rx / 2.0
            hh = float(s.get('height', 0)) * abs(float(s.get('scaleY', 1))) * ry / 2.0
            
            fx = -1 if (float(s.get('scaleX', 1)) < 0) ^ bool(s.get('flipX', False)) else 1
            fy = -1 if (float(s.get('scaleY', 1)) < 0) ^ bool(s.get('flipY', False)) else 1
            mat_raw = fitz.Matrix(fx, 0, 0, fy, 0, 0) * fitz.Matrix(float(s.get('angle', 0)))
            mat_transform = fitz.Matrix(1, 0, 0, 1, -center.x, -center.y) * mat_raw * fitz.Matrix(1, 0, 0, 1, center.x, center.y)

            if stype == 'rect':
                rp = [p * mat_transform for p in [fitz.Point(center_x - hw, center_y - hh), fitz.Point(center_x + hw, center_y - hh), fitz.Point(center_x - hw, center_y + hh), fitz.Point(center_x + hw, center_y + hh)]]
                page.draw_quad(fitz.Quad(rp[0], rp[1], rp[2], rp[3]), color=s_color, fill=f_color, width=s_width, fill_opacity=opacity, stroke_opacity=opacity)
            elif stype == 'triangle':
                rp = [p * mat_transform for p in [fitz.Point(center_x, center_y - hh), fitz.Point(center_x, center_y - hh), fitz.Point(center_x - hw, center_y + hh), fitz.Point(center_x + hw, center_y + hh)]]
                page.draw_quad(fitz.Quad(rp[0], rp[1], rp[2], rp[3]), color=s_color, fill=f_color, width=s_width, fill_opacity=opacity, stroke_opacity=opacity)
            elif stype == 'circle':
                page.draw_oval(fitz.Rect(center_x - hw, center_y - hh, center_x + hw, center_y + hh), color=s_color, fill=f_color, width=s_width, fill_opacity=opacity, stroke_opacity=opacity, morph=(center, mat_raw))
            elif stype == 'i-text':
                font_path = self.sys_fonts.get(s.get('fontFamily', 'helv'), None)
                fontsize = float(s.get('fontSize', 48)) * ry
                page.insert_text(fitz.Point(center_x - hw, center_y - hh + (fontsize * 0.8)), str(s.get('text', '')), fontsize=fontsize, fontname="korean" if font_path else "helv", fontfile=font_path, color=f_color if f_color else (0,0,0), fill_opacity=opacity, morph=(center, mat_raw))
            elif stype == 'image' and 'base64,' in str(s.get('src', '')):
                page.insert_image(fitz.Rect(center_x - hw, center_y - hh, center_x + hw, center_y + hh).quad * mat_transform, stream=base64.b64decode(s.get('src').split('base64,')[1]))

    def export_pages(self, pages_sequence, options):
        if not pages_sequence: return {"success": False, "error": "페이지 없음"}
        fmt = options.get('format', 'merge_pdf')
        try:
            if fmt == 'merge_pdf':
                save_result = self._window.create_file_dialog(webview.FileDialog.SAVE, save_filename="edited_document.pdf", file_types=('PDF 파일 (*.pdf)',))
                if not save_result: return {"success": False, "error": "취소"}
                save_path = save_result[0]
            else:
                folder_result = self._window.create_file_dialog(webview.FileDialog.FOLDER)
                if not folder_result: return {"success": False, "error": "취소"}
                output_folder = folder_result[0]

            if self._window:
                self._window.evaluate_js("let s=document.getElementById('ultimate-shield');if(!s){s=document.createElement('div');s.id='ultimate-shield';s.style.cssText='position:fixed;top:0;left:0;width:100vw;height:100vh;z-index:2147483647;background:transparent;cursor:wait;';document.body.appendChild(s);} showLoading('결과물 파일을 생성하고 있습니다...'); startProgressPolling();")
            self.update_progress(5, "1단계: 결과물 추출 준비 중...")

            self.update_progress(15, "2단계: 원본 파일 스캔 및 암호 해제 중...")
            sources_used = set(os.path.normcase(os.path.abspath(p['source'].replace('file:///', '').replace('file://', ''))) for p in pages_sequence)
            out_pdf = fitz.open()
            open_docs = {}

            if len(sources_used) == 1:
                single_src = list(sources_used)[0]
                if single_src in self.doc_cache:
                    out_pdf = fitz.open("pdf", self.doc_cache[single_src].tobytes())
                else:
                    out_pdf = fitz.open(single_src)
                    if out_pdf.is_encrypted:
                        if out_pdf.authenticate(""):
                            temp_bytes = out_pdf.tobytes(encryption=fitz.PDF_ENCRYPT_NONE)
                            out_pdf.close(); out_pdf = fitz.open("pdf", temp_bytes)
                        else:
                            out_pdf.close()
                            if self._window: self._window.evaluate_js("stopProgressPolling(); hideLoading();")
                            return {"success": False, "error": "암호 해제 실패"}
                
                self.update_progress(40, "3단계: 페이지 순서 조립 및 텍스트/도형 적용 중...")
                out_pdf.select([p['page_index'] for p in pages_sequence])
                for i, info in enumerate(pages_sequence):
                    page = out_pdf[i]
                    if info.get('rotation', 0): page.set_rotation((page.rotation + info.get('rotation', 0)) % 360)
                    if info.get('shapes'): self._draw_shapes_on_page(page, info.get('shapes'))
                    
                    if info.get('rasterOverlay'):
                        img_data = base64.b64decode(info.get('rasterOverlay').split('base64,')[1])
                        page.insert_image(page.rect, stream=img_data)
            else:
                chunks = []; current_chunk = None
                for page_info in pages_sequence:
                    src = os.path.normcase(os.path.abspath(page_info['source'].replace('file:///', '').replace('file://', '')))
                    page_info['clean_source'] = src
                    idx = page_info['page_index']
                    
                    if src not in open_docs:
                        if src in self.doc_cache:
                            open_docs[src] = fitz.open("pdf", self.doc_cache[src].tobytes())
                        else:
                            doc = fitz.open(src)
                            if doc.is_encrypted:
                                if doc.authenticate(""):
                                    temp_bytes = doc.tobytes(encryption=fitz.PDF_ENCRYPT_NONE)
                                    doc.close(); doc = fitz.open("pdf", temp_bytes)
                                else:
                                    doc.close(); 
                                    if self._window: self._window.evaluate_js("stopProgressPolling(); hideLoading();")
                                    return {"success": False, "error": "암호 해제 실패"}
                            open_docs[src] = doc

                    if current_chunk is None: current_chunk = {'src': src, 'start': idx, 'end': idx, 'infos': [page_info]}
                    else:
                        if current_chunk['src'] == src and idx == current_chunk['end'] + 1:
                            current_chunk['end'] = idx; current_chunk['infos'].append(page_info)
                        else:
                            chunks.append(current_chunk); current_chunk = {'src': src, 'start': idx, 'end': idx, 'infos': [page_info]}
                if current_chunk: chunks.append(current_chunk)

                current_out_idx = 0
                total_chunks = len(chunks)
                for chunk_idx, chunk in enumerate(chunks):
                    self.update_progress(40 + int((chunk_idx / total_chunks) * 30), f"3단계: 여러 파일 조립 중 ({chunk_idx + 1} / {total_chunks})")
                    out_pdf.insert_pdf(open_docs[chunk['src']], from_page=chunk['start'], to_page=chunk['end'])
                    for i, info in enumerate(chunk['infos']):
                        page = out_pdf[current_out_idx + i]
                        if info.get('rotation', 0): page.set_rotation((page.rotation + info.get('rotation', 0)) % 360)
                        if info.get('shapes'): self._draw_shapes_on_page(page, info.get('shapes'))
                        
                        if info.get('rasterOverlay'):
                            img_data = base64.b64decode(info.get('rasterOverlay').split('base64,')[1])
                            page.insert_image(page.rect, stream=img_data)
                            
                    current_out_idx += len(chunk['infos'])
                for doc in open_docs.values(): doc.close()

            if fmt == 'merge_pdf':
                self.update_progress(80, "4단계: 조합된 PDF 문서를 디스크에 저장 중...")
                # [해결 완료] 쓰레기 데이터(Garbage) 청소 및 압축(Deflate) 옵션 추가
                out_pdf.save(save_path, garbage=3, deflate=True, clean=True) 
                out_pdf.close()
                self.update_progress(100, "내보내기 완료!")
                return {"success": True, "saved_path": save_path}
            else:
                saved_count = 0
                total = len(out_pdf)
                for i in range(total):
                    self.update_progress(75 + int((i / total) * 20), f"4단계: 페이지 개별 추출 및 저장 중 ({i + 1} / {total})")
                    page = out_pdf[i]
                    if fmt == 'split_pdf':
                        temp_pdf = fitz.open()
                        temp_pdf.insert_pdf(out_pdf, from_page=i, to_page=i)
                        target_path = os.path.join(output_folder, f"page_{i+1:03d}.pdf")
                        if os.path.normcase(os.path.abspath(target_path)) in self.editor_loaded_files:
                            target_path = os.path.join(output_folder, f"page_{i+1:03d}_edited.pdf")
                        # [해결 완료] 쪼개서 저장할 때도 쓰레기 데이터 청소 및 압축 지시!
                        temp_pdf.save(target_path, garbage=3, deflate=True, clean=True)
                        temp_pdf.close()
                    elif fmt == 'split_image':
                        page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False).save(os.path.join(output_folder, f"page_{i+1:03d}.png"))
                    saved_count += 1
                out_pdf.close()
                self.update_progress(100, "내보내기 완료!")
                return {"success": True, "saved_path": f"{output_folder}\n(총 {saved_count}개 파일 저장)"}
        except Exception as e: return {"success": False, "error": f"저장 오류:\n{str(e)}"}

    # [신규] 편집된 문서를 조립하여 윈도우 인쇄 스풀러로 전송하는 전용 엔진
    def print_pages(self, pages_sequence):
        if not pages_sequence: return {"success": False, "error": "인쇄할 페이지가 없습니다."}
        try:
            if self._window:
                self._window.evaluate_js("let s=document.getElementById('ultimate-shield');if(!s){s=document.createElement('div');s.id='ultimate-shield';s.style.cssText='position:fixed;top:0;left:0;width:100vw;height:100vh;z-index:2147483647;background:transparent;cursor:wait;';document.body.appendChild(s);} showLoading('인쇄용 문서를 준비하고 있습니다...'); startProgressPolling();")
            
            self.update_progress(10, "1단계: 인쇄용 문서 병합 및 조립 준비 중...")

            self.update_progress(20, "2단계: 원본 파일 스캔 및 암호 해제 중...")
            sources_used = set(os.path.normcase(os.path.abspath(p['source'].replace('file:///', '').replace('file://', ''))) for p in pages_sequence)
            out_pdf = fitz.open()
            open_docs = {}

            if len(sources_used) == 1:
                single_src = list(sources_used)[0]
                if single_src in self.doc_cache:
                    out_pdf = fitz.open("pdf", self.doc_cache[single_src].tobytes())
                else:
                    out_pdf = fitz.open(single_src)
                    if out_pdf.is_encrypted:
                        if out_pdf.authenticate(""):
                            temp_bytes = out_pdf.tobytes(encryption=fitz.PDF_ENCRYPT_NONE)
                            out_pdf.close(); out_pdf = fitz.open("pdf", temp_bytes)
                        else:
                            out_pdf.close()
                            if self._window: self._window.evaluate_js("stopProgressPolling(); hideLoading();")
                            return {"success": False, "error": "암호 해제 실패"}
                
                self.update_progress(50, "3단계: 페이지 순서 배치 및 도장/텍스트 얹는 중...")
                out_pdf.select([p['page_index'] for p in pages_sequence])
                for i, info in enumerate(pages_sequence):
                    page = out_pdf[i]
                    if info.get('rotation', 0): page.set_rotation((page.rotation + info.get('rotation', 0)) % 360)
                    if info.get('shapes'): self._draw_shapes_on_page(page, info.get('shapes'))
                    
                    if info.get('rasterOverlay'):
                        img_data = base64.b64decode(info.get('rasterOverlay').split('base64,')[1])
                        page.insert_image(page.rect, stream=img_data)
            else:
                chunks = []; current_chunk = None
                for page_info in pages_sequence:
                    src = os.path.normcase(os.path.abspath(page_info['source'].replace('file:///', '').replace('file://', '')))
                    idx = page_info['page_index']
                    
                    if src not in open_docs:
                        if src in self.doc_cache:
                            open_docs[src] = fitz.open("pdf", self.doc_cache[src].tobytes())
                        else:
                            doc = fitz.open(src)
                            if doc.is_encrypted:
                                if doc.authenticate(""):
                                    temp_bytes = doc.tobytes(encryption=fitz.PDF_ENCRYPT_NONE)
                                    doc.close(); doc = fitz.open("pdf", temp_bytes)
                                else:
                                    doc.close(); 
                                    if self._window: self._window.evaluate_js("stopProgressPolling(); hideLoading();")
                                    return {"success": False, "error": "암호 해제 실패"}
                            open_docs[src] = doc

                    if current_chunk is None: current_chunk = {'src': src, 'start': idx, 'end': idx, 'infos': [page_info]}
                    else:
                        if current_chunk['src'] == src and idx == current_chunk['end'] + 1:
                            current_chunk['end'] = idx; current_chunk['infos'].append(page_info)
                        else:
                            chunks.append(current_chunk); current_chunk = {'src': src, 'start': idx, 'end': idx, 'infos': [page_info]}
                if current_chunk: chunks.append(current_chunk)

                current_out_idx = 0
                total_chunks = len(chunks)
                for chunk_idx, chunk in enumerate(chunks):
                    self.update_progress(40 + int((chunk_idx / total_chunks) * 40), f"3단계: 다중 파일 병합 조립 중 ({chunk_idx + 1} / {total_chunks})")
                    out_pdf.insert_pdf(open_docs[chunk['src']], from_page=chunk['start'], to_page=chunk['end'])
                    for i, info in enumerate(chunk['infos']):
                        page = out_pdf[current_out_idx + i]
                        if info.get('rotation', 0): page.set_rotation((page.rotation + info.get('rotation', 0)) % 360)
                        if info.get('shapes'): self._draw_shapes_on_page(page, info.get('shapes'))
                        
                        if info.get('rasterOverlay'):
                            img_data = base64.b64decode(info.get('rasterOverlay').split('base64,')[1])
                            page.insert_image(page.rect, stream=img_data)
                            
                    current_out_idx += len(chunk['infos'])
                for doc in open_docs.values(): doc.close()

            self.update_progress(90, "4단계: 임시 파일 생성 및 인쇄 스풀러 전송 중...")
            
            # 임시 폴더에 보안 처리된 인쇄용 PDF를 찍어냅니다.
            temp_path = os.path.join(tempfile.gettempdir(), f"print_{uuid.uuid4().hex}.pdf")
            out_pdf.save(temp_path, garbage=3, deflate=True, clean=True) 
            out_pdf.close()
            
            # 윈도우 OS의 기본 인쇄 다이얼로그 호출
            import platform
            if platform.system() == "Windows":
                os.startfile(temp_path, "print")
            
            self.update_progress(100, "프린터로 명령을 전송했습니다!")
            if self._window:
                self._window.evaluate_js("stopProgressPolling(); hideLoading(); let s=document.getElementById('ultimate-shield');if(s)s.remove();")
                
            return {"success": True}
        except Exception as e: 
            return {"success": False, "error": f"인쇄 오류:\n{str(e)}"}

    # =================================================================
    # 7. 보안 설정 제어 (저장/해제)
    # =================================================================
    def check_pdf_encrypted(self, file_path):
        try:
            doc = fitz.open(file_path)
            is_enc = doc.is_encrypted
            doc.close()
            return {"success": True, "is_encrypted": is_enc}
        except Exception as e: return {"success": False, "error": str(e)}

    def set_pdf_password(self, input_path, password):
        try:
            doc = fitz.open(input_path)
            if doc.is_encrypted: doc.close(); return {"success": False, "error": "이미 암호화된 파일입니다."}
            save_result = self._window.create_file_dialog(webview.FileDialog.SAVE, save_filename=os.path.basename(input_path.replace('.pdf', '_secured.pdf')), file_types=('PDF 파일 (*.pdf)',))
            if not save_result: return {"success": False, "error": "취소"}
            doc.save(save_result[0], encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw=password, user_pw=password)
            doc.close()
            return {"success": True, "saved_path": save_result[0]}
        except Exception as e: return {"success": False, "error": f"암호 설정 오류:\n{str(e)}"}

    def remove_pdf_password(self, input_path, password):
        try:
            doc = fitz.open(input_path)
            if not doc.is_encrypted: doc.close(); return {"success": False, "error": "암호화되지 않은 파일입니다."}
            if not doc.authenticate(password): doc.close(); return {"success": False, "error": "비밀번호가 올바르지 않습니다."}
            save_result = self._window.create_file_dialog(webview.FileDialog.SAVE, save_filename=os.path.basename(input_path.replace('.pdf', '_unlocked.pdf')), file_types=('PDF 파일 (*.pdf)',))
            if not save_result: return {"success": False, "error": "취소"}
            doc.save(save_result[0])
            doc.close()
            return {"success": True, "saved_path": save_result[0]}
        except Exception as e: return {"success": False, "error": f"암호 해제 오류:\n{str(e)}"}

    # =================================================================
    # 8. 프레젠테이션 (슬라이드쇼) 창 제어
    # =================================================================
    def get_screens(self):
        try: return {"success": True, "count": len(webview.screens) if hasattr(webview, 'screens') else 1}
        except Exception as e: return {"success": False, "error": str(e), "count": 1}

    # [수정] remote_pos 파라미터가 추가되었습니다.
    def open_presentation_window(self, pages_data, start_index, fullscreen, display_mode='current', crop_box=None, remote_pos='right'):
        if hasattr(self, '_pres_window') and self._pres_window:
            try: self._pres_window.destroy()
            except: pass
        # [핵심] 리모컨 위치 정보를 메모리에 저장하여 자바스크립트로 넘깁니다.
        self.pres_data = {"pages": pages_data, "index": start_index, "crop": crop_box, "remote_pos": remote_pos}
        html_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pdf_master.html')
        try:
            target_screen = None
            if hasattr(webview, 'screens') and display_mode == 'other' and len(webview.screens) > 1:
                target_screen = next((s for s in webview.screens if s != getattr(self._window, 'screen', None)), None)
            self._pres_window = webview.create_window('PenPDF Presentation', url=f'file://{html_file}', js_api=self, fullscreen=fullscreen, frameless=fullscreen, maximized=not fullscreen, easy_drag=False, on_top=fullscreen, screen=target_screen, hidden=True)
            return {"success": True}
        except Exception as e: return {"success": False, "error": str(e)}

    def show_presentation_window(self):
        if hasattr(self, '_pres_window') and self._pres_window:
            try: self._pres_window.show()
            except: pass
        return {"success": True}

    def consume_pres_data(self):
        data = getattr(self, 'pres_data', None)
        self.pres_data = None
        return data

    def close_presentation_window(self, last_index=None):
        if hasattr(self, '_pres_window') and self._pres_window:
            try: self._pres_window.destroy()
            except: pass
            self._pres_window = None
        if last_index is not None and self._window:
            try: self._window.evaluate_js(f"window.syncMainSelection({last_index})")
            except: pass
        return {"success": True}

    def sync_main_window(self, index):
        if self._window:
            try: self._window.evaluate_js(f"window.syncMainSelection({index})")
            except: pass
        return {"success": True}

    def sync_pres_window(self, index):
        if hasattr(self, '_pres_window') and self._pres_window:
            try:
                self._pres_window.evaluate_js(f"window.changeSlideFromMain({index})")
                if hasattr(self._pres_window, 'show'): self._pres_window.show()
            except: pass
        return {"success": True}

    def broadcast_crop_rect(self, crop_data):
        payload = json.dumps(crop_data) if crop_data else 'null'
        if self._window:
            try: self._window.evaluate_js(f"window.receiveCropRect({payload})")
            except: pass
        if hasattr(self, '_pres_window') and self._pres_window:
            try: self._pres_window.evaluate_js(f"window.receiveCropRect({payload})")
            except: pass
        return {"success": True}
    
    # =================================================================
    # [수정] 스마트 반 나누기 분석 엔진 (선택된 전체 페이지 교집합 여백 탐지)
    # =================================================================
    def analyze_split_pages(self, pages_data):
        try:
            # 자바스크립트에서 리스트 전체를 보내므로 리스트 처리 준비
            if not isinstance(pages_data, list):
                pages_data = [pages_data]
                
            direction = 'vertical'
            
            # [핵심] 모든 페이지를 스캔하며 '가장 좁은 여백(가장 바깥쪽 글씨)'을 찾기 위해 초깃값을 넉넉하게 잡습니다.
            min_margin_top = 30
            min_margin_bottom = 30
            min_margin_left = 30
            min_margin_right = 30
            
            first_page_checked = False
            
            for p in pages_data:
                src = os.path.normcase(os.path.abspath(p['source'].replace('file:///', '').replace('file://', '')))
                idx = p['page_index']
                
                if src in self.doc_cache:
                    doc = fitz.open("pdf", self.doc_cache[src].tobytes())
                else:
                    doc = fitz.open(src)
                    if doc.is_encrypted:
                        if doc.authenticate(""):
                            doc = fitz.open("pdf", doc.tobytes(encryption=fitz.PDF_ENCRYPT_NONE))
                        else:
                            continue
                
                page = doc.load_page(idx)
                w, h = page.rect.width, page.rect.height
                
                # 자르는 방향은 첫 번째 페이지의 가로세로 비율 기준으로만 잡습니다.
                if not first_page_checked:
                    direction = 'vertical' if w > h else 'horizontal'
                    first_page_checked = True
                
                blocks = page.get_text("blocks")
                if blocks:
                    # 현재 페이지 내용물의 상하좌우 실제 경계선(Bounding Box) 탐지
                    curr_min_x = min([b[0] for b in blocks if b[0] >= 0] + [w])
                    curr_min_y = min([b[1] for b in blocks if b[1] >= 0] + [h])
                    curr_max_x = max([b[2] for b in blocks if b[2] <= w] + [0])
                    curr_max_y = max([b[3] for b in blocks if b[3] <= h] + [0])
                    
                    # 내용물이 페이지 모서리에서 1/3 지점보다 바깥쪽에 있을 때만 진짜 여백으로 인정
                    curr_left = max(10, int(curr_min_x)) if curr_min_x < w / 3 else 30
                    curr_top = max(10, int(curr_min_y)) if curr_min_y < h / 3 else 30
                    curr_right = max(10, int(w - curr_max_x)) if (w - curr_max_x) < w / 3 else 30
                    curr_bottom = max(10, int(h - curr_max_y)) if (h - curr_max_y) < h / 3 else 30
                    
                    # [비교 연산] 기존에 찾은 여백보다 현재 페이지의 여백이 더 좁으면(=글씨가 더 튀어나왔으면) 갱신
                    min_margin_left = min(min_margin_left, curr_left)
                    min_margin_top = min(min_margin_top, curr_top)
                    min_margin_right = min(min_margin_right, curr_right)
                    min_margin_bottom = min(min_margin_bottom, curr_bottom)
                    
            margin = {
                "top": min_margin_top,
                "bottom": min_margin_bottom,
                "left": min_margin_left,
                "right": min_margin_right
            }
                
            return {"success": True, "data": {"direction": direction, "margins": margin}}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # =================================================================
    # [수정] 목표 용지 크기 및 여백 적용 스마트 분할 엔진 (3/4분할, 자동 방향 제거)
    # =================================================================
    def split_pages_advanced(self, pages_data, options):
        try:
            results = []
            target_size = options.get('size', 'A4')
            direction = options.get('direction', 'vertical')
            split_count = str(options.get('split_count', '2'))
            orientation = options.get('orientation', 'portrait') # [수정] 기본값 portrait으로 설정
            
            m_top = int(options.get('margins', {}).get('top', 30))
            m_bottom = int(options.get('margins', {}).get('bottom', 30))
            m_left = int(options.get('margins', {}).get('left', 30))
            m_right = int(options.get('margins', {}).get('right', 30))

            paper_rect = fitz.paper_rect(target_size)
            pw, ph = paper_rect.width, paper_rect.height

            # [수정] '자동' 옵션 제거. 무조건 사용자가 선택한(또는 초기 세팅된) 가로/세로 방향으로 강제 고정합니다.
            if orientation == 'landscape':
                if pw < ph: 
                    pw, ph = ph, pw
            elif orientation == 'portrait':
                if pw > ph: 
                    pw, ph = ph, pw
                
            safe_rect = fitz.Rect(m_left, m_top, pw - m_right, ph - m_bottom)

            for p in pages_data:
                base_name = os.path.splitext(p.get('filename', 'split'))[0]
                
                src = os.path.normcase(os.path.abspath(p['source'].replace('file:///', '').replace('file://', '')))
                idx = p['page_index']
                
                if src in self.doc_cache:
                    doc = fitz.open("pdf", self.doc_cache[src].tobytes())
                else:
                    doc = fitz.open(src)
                    if doc.is_encrypted:
                        if doc.authenticate(""):
                            doc = fitz.open("pdf", doc.tobytes(encryption=fitz.PDF_ENCRYPT_NONE))
                        else:
                            continue
                
                page = doc.load_page(idx)
                w, h = page.rect.width, page.rect.height
                
                rects = []
                if split_count == '4':
                    rects = [
                        fitz.Rect(0, 0, w/2, h/2),
                        fitz.Rect(w/2, 0, w, h/2),
                        fitz.Rect(0, h/2, w/2, h),
                        fitz.Rect(w/2, h/2, w, h)
                    ]
                elif split_count == '3':
                    if direction == 'vertical':
                        rects = [
                            fitz.Rect(0, 0, w/3, h),
                            fitz.Rect(w/3, 0, w*2/3, h),
                            fitz.Rect(w*2/3, 0, w, h)
                        ]
                    else:
                        rects = [
                            fitz.Rect(0, 0, w, h/3),
                            fitz.Rect(0, h/3, w, h*2/3),
                            fitz.Rect(0, h*2/3, w, h)
                        ]
                else:
                    if direction == 'vertical':
                        rects = [fitz.Rect(0, 0, w/2, h), fitz.Rect(w/2, 0, w, h)]
                    else:
                        rects = [fitz.Rect(0, 0, w, h/2), fitz.Rect(0, h/2, w, h)]
                
                new_doc = fitz.open()
                
                for i, r_clip in enumerate(rects):
                    pg = new_doc.new_page(width=pw, height=ph)
                    pg.show_pdf_page(safe_rect, doc, idx, clip=r_clip, keep_proportion=True)
                
                temp_path = os.path.join(tempfile.gettempdir(), f"split_{uuid.uuid4().hex}.pdf")
                new_doc.save(temp_path)
                new_doc.close()
                doc.close()
                
                self.editor_loaded_files.add(temp_path)
                
                new_pages_info = []
                for i in range(len(rects)):
                    thumb = self.get_page_thumbnail(temp_path, i)['image']
                    new_pages_info.append({
                        'id': str(uuid.uuid4()), 
                        'source': temp_path, 
                        'filename': f"{base_name}_{i+1}", 
                        'page_index': i, 
                        'thumbnail': thumb
                    })
                
                results.append({
                    'original_id': p['id'],
                    'new_pages': new_pages_info
                })
            
            return {"success": True, "results": results}
        except Exception as e:
            return {"success": False, "error": str(e)}

if __name__ == '__main__':
    api = PdfApi()
    current_dir = os.path.dirname(os.path.abspath(__file__))
    html_file = os.path.join(current_dir, 'pdf_master.html')
    
    # [수정] 창의 최소 크기를 (1100, 700)으로 강제 제한하여 UI 찌그러짐을 원천 차단!
    window = webview.create_window('PenPDF', url=f'file://{html_file}', js_api=api, width=1400, height=800, min_size=(1100, 700), maximized=False)
    api.set_window(window)

    def on_closing():
        if api.force_close: return True
        def ask_js(): window.evaluate_js('checkAppExit()')
        threading.Timer(0.1, ask_js).start()
        return False

    window.events.closing += on_closing
    webview.start(debug=False, gui='edgechromium')