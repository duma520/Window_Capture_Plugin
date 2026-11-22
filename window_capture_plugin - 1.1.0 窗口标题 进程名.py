import cv2
import numpy as np
import win32gui
import win32ui
import win32con
import ctypes
import comtypes
from typing import Optional, Dict, List, Union, Tuple
import logging
from enum import Enum
import json
from fastapi import FastAPI, HTTPException
import uvicorn
import argparse
from PIL import ImageGrab
import dxcam
import mss
from pywincffi import graphicscapture
import uiautomation as auto
import psutil
import win32process

class CaptureMethod(Enum):
    """支持的捕获方法枚举"""
    BITBLT = 1
    DESKTOP_DUPLICATION = 2
    MAGNIFICATION_API = 3
    UIAUTOMATION = 4
    PIL_FALLBACK = 5
    DXGI = 6
    MSS = 7
    GRAPHICS_CAPTURE = 8
    OPENCV = 9
    CLIENT_AREA = 10

class WindowCapturePlugin:
    def __init__(self, config: Optional[Dict] = None):
        """
        初始化窗口捕获插件
        
        :param config: 配置字典，包含各种参数
        """
        self.config = config or {
            'default_method': CaptureMethod.BITBLT,
            'enable_preprocess': True,
            'low_performance_mode': False,
            'enable_color': True,
            'log_level': logging.INFO
        }
        
        # 设置日志
        logging.basicConfig(
            level=self.config.get('log_level', logging.INFO),
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger('WindowCapturePlugin')
        
        # 初始化DXCam
        self.dxcam_camera = None
        try:
            self.dxcam_camera = dxcam.create()
        except Exception as e:
            self.logger.warning(f"Failed to initialize DXCam: {e}")
    
    def find_window_by_title_or_process(
        self,
        window_title: Optional[str] = None,
        process_name: Optional[str] = None,
        exact_match: bool = True
    ) -> Optional[int]:
        """
        通过窗口标题或进程名查找窗口句柄
        
        :param window_title: 窗口标题（可选）
        :param process_name: 进程名（可选，如"notepad.exe"）
        :param exact_match: 是否精确匹配标题
        :return: 窗口句柄，找不到返回None
        """
        def enum_windows_callback(hwnd, params):
            title_match = False
            process_match = False
            
            # 检查窗口标题
            if window_title is not None:
                title = win32gui.GetWindowText(hwnd)
                if exact_match:
                    if title == window_title:
                        title_match = True
                else:
                    if window_title.lower() in title.lower():
                        title_match = True
            
            # 检查进程名
            if process_name is not None:
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    process = psutil.Process(pid)
                    current_process_name = process.name().lower()
                    if process_name.lower() in current_process_name:
                        process_match = True
                except:
                    pass
            
            # 根据参数决定匹配逻辑
            if window_title is not None and process_name is not None:
                if title_match and process_match:
                    params.append(hwnd)
            elif window_title is not None:
                if title_match:
                    params.append(hwnd)
            elif process_name is not None:
                if process_match:
                    params.append(hwnd)
            
            return True
        
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        if windows:
            return windows[0]  # 返回第一个匹配的窗口
        return None
    
    def capture_window(
        self,
        window_title: Optional[str] = None,
        process_name: Optional[str] = None,
        method: Optional[CaptureMethod] = None,
        region: Optional[Tuple[int, int, int, int]] = None,
        preprocess: Optional[bool] = None,
        exact_match: bool = True
    ) -> Optional[np.ndarray]:
        """
        捕获指定窗口
        
        :param window_title: 窗口标题（可选）
        :param process_name: 进程名（可选，如"notepad.exe"）
        :param method: 捕获方法，如果为None则使用默认方法
        :param region: 捕获区域 (x, y, width, height)，相对于窗口
        :param preprocess: 是否预处理图像，None则使用配置中的设置
        :param exact_match: 是否精确匹配标题
        :return: 捕获的图像 (BGR格式)，失败返回None
        """
        if window_title is None and process_name is None:
            raise ValueError("Either window_title or process_name must be provided")
        
        method = method or self.config['default_method']
        preprocess = preprocess if preprocess is not None else self.config['enable_preprocess']
        
        try:
            hwnd = self.find_window_by_title_or_process(window_title, process_name, exact_match)
            if not hwnd:
                self.logger.error(f"Window not found (title: {window_title}, process: {process_name})")
                return None
            
            img = None
            
            # 根据选择的方法调用对应的捕获函数
            if method == CaptureMethod.BITBLT:
                img = self._capture_bitblt(hwnd)
            elif method == CaptureMethod.DESKTOP_DUPLICATION:
                img = self._capture_desktop_duplication(hwnd)
            elif method == CaptureMethod.MAGNIFICATION_API:
                img = self._capture_magnification(hwnd)
            elif method == CaptureMethod.UIAUTOMATION:
                img = self._capture_uiautomation(hwnd)
            elif method == CaptureMethod.PIL_FALLBACK:
                img = self._capture_pil_fallback(hwnd)
            elif method == CaptureMethod.DXGI:
                img = self._capture_dxgi(hwnd, region)
            elif method == CaptureMethod.MSS:
                img = self._capture_mss(hwnd, region)
            elif method == CaptureMethod.GRAPHICS_CAPTURE:
                img = self._capture_graphics_capture(hwnd)
            elif method == CaptureMethod.OPENCV:
                img = self._capture_opencv(hwnd)
            elif method == CaptureMethod.CLIENT_AREA:
                img = self._capture_client_area(hwnd)
            
            if img is None:
                self.logger.warning(f"Method {method.name} failed, trying fallback methods")
                img = self._try_all_methods(hwnd, region)
            
            if img is not None and preprocess:
                img = self._preprocess_image(img)
            
            return img
            
        except Exception as e:
            self.logger.error(f"Error capturing window (title: {window_title}, process: {process_name}): {e}", exc_info=True)
            return None
    
    def _try_all_methods(self, hwnd: int, region: Optional[Tuple[int, int, int, int]] = None) -> Optional[np.ndarray]:
        """尝试所有可用的捕获方法"""
        methods = [
            CaptureMethod.BITBLT,
            CaptureMethod.DESKTOP_DUPLICATION,
            CaptureMethod.MAGNIFICATION_API,
            CaptureMethod.UIAUTOMATION,
            CaptureMethod.PIL_FALLBACK,
            CaptureMethod.DXGI,
            CaptureMethod.MSS,
            CaptureMethod.GRAPHICS_CAPTURE,
            CaptureMethod.CLIENT_AREA
        ]
        
        for method in methods:
            try:
                if method == CaptureMethod.BITBLT:
                    img = self._capture_bitblt(hwnd)
                elif method == CaptureMethod.DESKTOP_DUPLICATION:
                    img = self._capture_desktop_duplication(hwnd)
                elif method == CaptureMethod.MAGNIFICATION_API:
                    img = self._capture_magnification(hwnd)
                elif method == CaptureMethod.UIAUTOMATION:
                    img = self._capture_uiautomation(hwnd)
                elif method == CaptureMethod.PIL_FALLBACK:
                    img = self._capture_pil_fallback(hwnd)
                elif method == CaptureMethod.DXGI:
                    img = self._capture_dxgi(hwnd, region)
                elif method == CaptureMethod.MSS:
                    img = self._capture_mss(hwnd, region)
                elif method == CaptureMethod.GRAPHICS_CAPTURE:
                    img = self._capture_graphics_capture(hwnd)
                elif method == CaptureMethod.CLIENT_AREA:
                    img = self._capture_client_area(hwnd)
                
                if img is not None:
                    self.logger.info(f"Fallback method {method.name} succeeded")
                    return img
                
            except Exception as e:
                self.logger.debug(f"Method {method.name} failed: {e}")
        
        return None
    
    def _capture_bitblt(self, hwnd: int) -> Optional[np.ndarray]:
        """方法1: 使用BitBlt捕获窗口"""
        try:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top
            
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()
            
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            
            saveDC.SelectObject(saveBitMap)
            result = saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)
            
            if result:
                bmpinfo = saveBitMap.GetInfo()
                bmpstr = saveBitMap.GetBitmapBits(True)
                img = np.frombuffer(bmpstr, dtype='uint8').reshape((bmpinfo['bmHeight'], bmpinfo['bmWidth'], 4))
                img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
                
                win32gui.DeleteObject(saveBitMap.GetHandle())
                saveDC.DeleteDC()
                mfcDC.DeleteDC()
                win32gui.ReleaseDC(hwnd, hwndDC)
                
                return img
                
        except Exception as e:
            self.logger.debug(f"BitBlt method failed: {e}")
            return None
    
    def _capture_desktop_duplication(self, hwnd: int) -> Optional[np.ndarray]:
        """方法2: 使用Desktop Duplication API捕获窗口"""
        try:
            comtypes.CoInitialize()
            
            D3D11CreateDevice = ctypes.windll.d3d11.D3D11CreateDevice
            DXGI_OUTDUPL_FRAME_INFO = ctypes.c_void_p
            DXGI_OUTDUPL_DESC = ctypes.c_void_p
            
            device = ctypes.c_void_p()
            hr = D3D11CreateDevice(
                None, 
                ctypes.c_int(1),
                None, 
                0, 
                None, 
                0, 
                ctypes.c_int(7),
                ctypes.byref(device), 
                None, 
                None
            )
            if hr != 0:
                raise Exception(f"D3D11CreateDevice failed with HRESULT: {hr}")
            
            dxgi_device = ctypes.POINTER(ctypes.c_void_p)()
            hr = device.QueryInterface(
                comtypes.GUID("{54ec77fa-1377-44e6-8c32-88fd5f44c84c}"), 
                ctypes.byref(dxgi_device)
            )
            
            adapter = ctypes.POINTER(ctypes.c_void_p)()
            hr = dxgi_device.GetAdapter(ctypes.byref(adapter))
            
            output = ctypes.POINTER(ctypes.c_void_p)()
            hr = adapter.EnumOutputs(0, ctypes.byref(output))
            
            desc = DXGI_OUTDUPL_DESC()
            output.GetDesc(ctypes.byref(desc))
            
            duplication = ctypes.POINTER(ctypes.c_void_p)()
            hr = output.DuplicateOutput(device, ctypes.byref(duplication))
            
            frame_info = DXGI_OUTDUPL_FRAME_INFO()
            desktop_resource = ctypes.POINTER(ctypes.c_void_p)()
            hr = duplication.AcquireNextFrame(1000, ctypes.byref(frame_info), ctypes.byref(desktop_resource))
            
            texture = ctypes.POINTER(ctypes.c_void_p)()
            hr = desktop_resource.QueryInterface(
                comtypes.GUID("{6f15aaf2-d208-4e89-9ab4-489535d34f9c}"), 
                ctypes.byref(texture))
            )
            
            mapped_rect = ctypes.c_void_p()
            hr = texture.Map(0, ctypes.byref(mapped_rect))
            
            img = np.frombuffer(
                mapped_rect.pBits, 
                dtype=np.uint8, 
                count=desc.ModeDesc.Width * desc.ModeDesc.Height * 4
            ).reshape((desc.ModeDesc.Height, desc.ModeDesc.Width, 4))
            
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
            
            texture.Unmap(0)
            duplication.ReleaseFrame()
            comtypes.CoUninitialize()
            
            return img
            
        except Exception as e:
            self.logger.debug(f"Desktop Duplication API method failed: {e}")
            return None
    
    def _capture_magnification(self, hwnd: int) -> Optional[np.ndarray]:
        """方法3: 使用Magnification API捕获窗口"""
        try:
            if not ctypes.windll.magnification.MagInitialize():
                raise Exception("MagInitialize failed")
            
            hwndMag = ctypes.windll.user32.CreateWindowExA(
                0, "Magnifier", "MagnifierWindow", 
                ctypes.wintypes.DWORD(0x80000000), 
                0, 0, 0, 0, 
                None, None, None, None
            )
            
            if not hwndMag:
                raise Exception("CreateWindowEx failed")
            
            source_rect = ctypes.wintypes.RECT()
            win32gui.GetWindowRect(hwnd, ctypes.byref(source_rect))
            
            ctypes.windll.user32.SetWindowPos(
                hwndMag, None, 
                source_rect.left, source_rect.top, 
                source_rect.right - source_rect.left, 
                source_rect.bottom - source_rect.top, 
                0
            )
            
            if not ctypes.windll.magnification.MagSetWindowSource(
                hwndMag, source_rect
            ):
                raise Exception("MagSetWindowSource failed")
            
            hdc = ctypes.windll.user32.GetDC(hwndMag)
            if not hdc:
                raise Exception("GetDC failed")
            
            hdcMem = ctypes.windll.gdi32.CreateCompatibleDC(hdc)
            if not hdcMem:
                raise Exception("CreateCompatibleDC failed")
            
            hBitmap = ctypes.windll.gdi32.CreateCompatibleBitmap(
                hdc, 
                source_rect.right - source_rect.left, 
                source_rect.bottom - source_rect.top
            )
            if not hBitmap:
                raise Exception("CreateCompatibleBitmap failed")
            
            hOld = ctypes.windll.gdi32.SelectObject(hdcMem, hBitmap)
            
            if not ctypes.windll.gdi32.BitBlt(
                hdcMem, 0, 0, 
                source_rect.right - source_rect.left, 
                source_rect.bottom - source_rect.top, 
                hdc, 0, 0, 
                0x00CC0020  # SRCCOPY
            ):
                raise Exception("BitBlt failed")
            
            bmpinfo = win32ui.CreateBitmapInfo()
            bmpinfo.bmiHeader.biSize = ctypes.sizeof(bmpinfo.bmiHeader)
            bmpinfo.bmiHeader.biWidth = source_rect.right - source_rect.left
            bmpinfo.bmiHeader.biHeight = -(source_rect.bottom - source_rect.top)
            bmpinfo.bmiHeader.biPlanes = 1
            bmpinfo.bmiHeader.biBitCount = 32
            bmpinfo.bmiHeader.biCompression = 0
            
            buffer_size = (source_rect.right - source_rect.left) * (source_rect.bottom - source_rect.top) * 4
            buffer = (ctypes.c_ubyte * buffer_size)()
            
            if ctypes.windll.gdi32.GetDIBits(
                hdcMem, hBitmap, 0, 
                source_rect.bottom - source_rect.top, 
                buffer, 
                ctypes.byref(bmpinfo), 
                0
            ) == 0:
                raise Exception("GetDIBits failed")
            
            img = np.frombuffer(buffer, dtype=np.uint8).reshape(
                (source_rect.bottom - source_rect.top, 
                source_rect.right - source_rect.left, 4)
            )
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
            
            ctypes.windll.gdi32.SelectObject(hdcMem, hOld)
            ctypes.windll.gdi32.DeleteObject(hBitmap)
            ctypes.windll.gdi32.DeleteDC(hdcMem)
            ctypes.windll.user32.ReleaseDC(hwndMag, hdc)
            ctypes.windll.magnification.MagUninitialize()
            
            return img
            
        except Exception as e:
            self.logger.debug(f"Magnification API method failed: {e}")
            return None
    
    def _capture_uiautomation(self, hwnd: int) -> Optional[np.ndarray]:
        """方法4: 使用UIAutomation API捕获窗口"""
        try:
            window_title = win32gui.GetWindowText(hwnd)
            window_element = auto.WindowControl(searchDepth=1, Name=window_title)
            if not window_element.Exists():
                raise Exception("Window not found")
            
            rect = window_element.BoundingRectangle
            left, top, right, bottom = rect.left, rect.top, rect.right, rect.bottom
            
            from PIL import ImageGrab
            screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
            return cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
        except Exception as e:
            self.logger.debug(f"UIAutomation API method failed: {e}")
            return None
    
    def _capture_pil_fallback(self, hwnd: int) -> Optional[np.ndarray]:
        """方法5: 使用PIL截图作为回退方案"""
        try:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            from PIL import ImageGrab
            screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
            return cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
        except Exception as e:
            self.logger.debug(f"PIL fallback method failed: {e}")
            return None
    
    def _capture_dxgi(self, hwnd: int, region: Optional[Tuple[int, int, int, int]] = None) -> Optional[np.ndarray]:
        """方法6: 使用DXGI/DirectX捕获"""
        try:
            if self.dxcam_camera is None:
                raise Exception("DXCam not initialized")
            
            rect = win32gui.GetWindowRect(hwnd)
            region_params = (rect[0], rect[1], rect[2]-rect[0], rect[3]-rect[1])
            
            if region:
                region_params = (
                    rect[0] + region[0],
                    rect[1] + region[1],
                    region[2],
                    region[3]
                )
            
            img = self.dxcam_camera.grab(region=region_params)
            if img is not None:
                return img
            raise Exception("DXGI capture returned None")
            
        except Exception as e:
            self.logger.debug(f"DXGI method failed: {e}")
            return None
    
    def _capture_mss(self, hwnd: int, region: Optional[Tuple[int, int, int, int]] = None) -> Optional[np.ndarray]:
        """方法7: 使用MSS跨平台捕获"""
        try:
            with mss.mss() as sct:
                rect = win32gui.GetWindowRect(hwnd)
                monitor = {
                    "top": rect[1], 
                    "left": rect[0], 
                    "width": rect[2]-rect[0], 
                    "height": rect[3]-rect[1]
                }
                
                if region:
                    monitor = {
                        "top": monitor["top"] + region[1],
                        "left": monitor["left"] + region[0],
                        "width": region[2],
                        "height": region[3]
                    }
                
                img = np.array(sct.grab(monitor))
                return cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
                
        except Exception as e:
            self.logger.debug(f"MSS method failed: {e}")
            return None
    
    def _capture_graphics_capture(self, hwnd: int) -> Optional[np.ndarray]:
        """方法8: 使用Windows GraphicsCapture API (Win10+)"""
        try:
            item = graphicscapture.Capture.create_for_window(hwnd)
            frame = item.capture_frame()
            img = np.array(frame)
            return cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
            
        except Exception as e:
            self.logger.debug(f"GraphicsCapture method failed: {e}")
            return None
    
    def _capture_opencv(self, hwnd: int) -> Optional[np.ndarray]:
        """方法9: 使用OpenCV截图"""
        try:
            rect = win32gui.GetWindowRect(hwnd)
            with mss.mss() as sct:
                monitor = {
                    "top": rect[1], 
                    "left": rect[0], 
                    "width": rect[2]-rect[0], 
                    "height": rect[3]-rect[1]
                }
                img = np.array(sct.grab(monitor))
                return cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
                
        except Exception as e:
            self.logger.debug(f"OpenCV method failed: {e}")
            return None
    
    def _capture_client_area(self, hwnd: int) -> Optional[np.ndarray]:
        """方法10: 捕获窗口客户区"""
        try:
            try:
                left, top, right, bottom = win32gui.GetClientRect(hwnd)
                if right == 0 or bottom == 0:
                    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                    right -= left
                    bottom -= top
                    left = 0
                    top = 0
            except:
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                right -= left
                bottom -= top
                left = 0
                top = 0

            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()

            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, right, bottom)

            saveDC.SelectObject(saveBitMap)
            saveDC.BitBlt((0, 0), (right, bottom), mfcDC, (0, 0), win32con.SRCCOPY)

            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)

            img = np.frombuffer(bmpstr, dtype='uint8').reshape((bmpinfo['bmHeight'], bmpinfo['bmWidth'], 4))
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)

            return img
            
        except Exception as e:
            self.logger.debug(f"Client area method failed: {e}")
            return None
    
    def _preprocess_image(self, img: np.ndarray) -> np.ndarray:
        """图像预处理"""
        if self.config.get('low_performance_mode', False) or img.shape[1] > 1280:
            # 降采样 (保持宽高比)
            if max(img.shape[:2]) > 1280:
                scale = 1280 / max(img.shape[:2])
                img = cv2.resize(img, (0,0), fx=scale, fy=scale)
            
            # 智能灰度化
            if not self.config.get('enable_color', True):
                if img.ndim == 3:
                    gray_mean = cv2.mean(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY))[0]
                    if gray_mean > 40:
                        img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
            
            # 限制对比度自适应直方图均衡化
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
            lab = cv2.cvtColor(img, cv2.COLOR_BGR2LAB)
            lab[...,0] = clahe.apply(lab[...,0])
            img = cv2.cvtColor(lab, cv2.COLOR_LAB2BGR)
        
        return img
    
    def list_windows(self) -> List[Dict[str, Union[str, int]]]:
        """列出所有可见窗口"""
        windows = []
        
        def enum_windows_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:
                    try:
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        process = psutil.Process(pid)
                        process_name = process.name()
                    except:
                        process_name = "Unknown"
                    
                    windows.append({
                        'hwnd': hwnd,
                        'title': title,
                        'process': process_name
                    })
            return True
        
        win32gui.EnumWindows(enum_windows_callback, None)
        return windows

# REST API 服务
app = FastAPI()
capture_plugin = WindowCapturePlugin()

@app.post("/capture")
async def capture_window_api(
    window_title: Optional[str] = None,
    process_name: Optional[str] = None,
    method: Optional[str] = None,
    region: Optional[str] = None,
    preprocess: Optional[bool] = None,
    exact_match: bool = True
):
    try:
        if window_title is None and process_name is None:
            raise ValueError("Either window_title or process_name must be provided")
        
        method_enum = CaptureMethod[method.upper()] if method else None
        
        region_tuple = None
        if region:
            region_tuple = tuple(map(int, region.split(',')))
            if len(region_tuple) != 4:
                raise ValueError("Region must be 4 integers: x,y,width,height")
        
        img = capture_plugin.capture_window(
            window_title=window_title,
            process_name=process_name,
            method=method_enum,
            region=region_tuple,
            preprocess=preprocess,
            exact_match=exact_match
        )
        
        if img is None:
            raise HTTPException(status_code=404, detail="Window capture failed")
        
        _, buffer = cv2.imencode('.png', img)
        return {"image": buffer.tobytes().hex()}
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/windows")
async def list_windows_api():
    try:
        windows = capture_plugin.list_windows()
        return {"windows": windows}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

def run_api_server(host: str = "0.0.0.0", port: int = 8000):
    """启动REST API服务器"""
    uvicorn.run(app, host=host, port=port)

# 命令行接口
def main():
    parser = argparse.ArgumentParser(description="Window Capture Plugin")
    subparsers = parser.add_subparsers(dest='command', required=True)
    
    # 捕获命令
    capture_parser = subparsers.add_parser('capture', help='Capture a window')
    capture_parser.add_argument('--window_title', help='Title of the window to capture')
    capture_parser.add_argument('--process_name', help='Process name of the window to capture (e.g. notepad.exe)')
    capture_parser.add_argument('--method', help='Capture method to use', 
                              choices=[m.name for m in CaptureMethod])
    capture_parser.add_argument('--region', help='Capture region (x,y,width,height)')
    capture_parser.add_argument('--output', help='Output image file path', required=True)
    capture_parser.add_argument('--preprocess', help='Enable image preprocessing', 
                              action='store_true')
    capture_parser.add_argument('--fuzzy_match', help='Use fuzzy match for window title', 
                              action='store_false', dest='exact_match')
    
    # 列表命令
    list_parser = subparsers.add_parser('list', help='List all visible windows')
    
    # API服务命令
    api_parser = subparsers.add_parser('serve', help='Start REST API server')
    api_parser.add_argument('--host', help='Host to bind to', default='0.0.0.0')
    api_parser.add_argument('--port', help='Port to listen on', type=int, default=8000)
    
    args = parser.parse_args()
    
    if args.command == 'capture':
        if args.window_title is None and args.process_name is None:
            print("Error: Either --window_title or --process_name must be provided")
            return
        
        region_tuple = None
        if args.region:
            region_tuple = tuple(map(int, args.region.split(',')))
            if len(region_tuple) != 4:
                print("Error: Region must be 4 integers: x,y,width,height")
                return
        
        method = CaptureMethod[args.method.upper()] if args.method else None
        img = capture_plugin.capture_window(
            window_title=args.window_title,
            process_name=args.process_name,
            method=method,
            region=region_tuple,
            preprocess=args.preprocess,
            exact_match=args.exact_match
        )
        
        if img is not None:
            cv2.imwrite(args.output, img)
            print(f"Window captured and saved to {args.output}")
        else:
            print("Failed to capture window")
    
    elif args.command == 'list':
        windows = capture_plugin.list_windows()
        for window in windows:
            print(f"HWND: {window['hwnd']}, Title: {window['title']}, Process: {window['process']}")
    
    elif args.command == 'serve':
        print(f"Starting API server on {args.host}:{args.port}")
        run_api_server(host=args.host, port=args.port)

if __name__ == "__main__":
    main()