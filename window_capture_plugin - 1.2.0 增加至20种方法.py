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
import win32api
import win32com.client
import pythoncom
import threading
import time

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
    WGC_API = 11
    DWM_API = 12
    GDIPLUS = 13
    DIRECTX_CAPTURE = 14
    WINDOWS_SCREENSHOT = 15
    DXGI_SURFACE_SHARING = 16
    WINDOWS_MEDIA_CAPTURE = 17
    D3D9_SURFACE_CAPTURE = 18
    WINDOWS_AI_VISION = 19
    APPLICATION_FRAME_HOST = 20

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
        
        # 初始化D3D9
        self.d3d9 = None
        try:
            self.d3d9 = ctypes.windll.d3d9
        except:
            self.logger.warning("Failed to load d3d9.dll")
    
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
            elif method == CaptureMethod.WGC_API:
                img = self._capture_wgc_api(hwnd)
            elif method == CaptureMethod.DWM_API:
                img = self._capture_dwm_api(hwnd)
            elif method == CaptureMethod.GDIPLUS:
                img = self._capture_gdiplus(hwnd)
            elif method == CaptureMethod.DIRECTX_CAPTURE:
                img = self._capture_directx(hwnd)
            elif method == CaptureMethod.WINDOWS_SCREENSHOT:
                img = self._capture_windows_screenshot(hwnd)
            elif method == CaptureMethod.DXGI_SURFACE_SHARING:
                img = self._capture_dxgi_surface_sharing(hwnd)
            elif method == CaptureMethod.WINDOWS_MEDIA_CAPTURE:
                img = self._capture_windows_media(hwnd)
            elif method == CaptureMethod.D3D9_SURFACE_CAPTURE:
                img = self._capture_d3d9_surface(hwnd)
            elif method == CaptureMethod.WINDOWS_AI_VISION:
                img = self._capture_windows_ai_vision(hwnd)
            elif method == CaptureMethod.APPLICATION_FRAME_HOST:
                img = self._capture_application_frame_host(hwnd)
            
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
            CaptureMethod.CLIENT_AREA,
            CaptureMethod.WGC_API,
            CaptureMethod.DWM_API,
            CaptureMethod.GDIPLUS,
            CaptureMethod.DIRECTX_CAPTURE,
            CaptureMethod.WINDOWS_SCREENSHOT,
            CaptureMethod.DXGI_SURFACE_SHARING,
            CaptureMethod.WINDOWS_MEDIA_CAPTURE,
            CaptureMethod.D3D9_SURFACE_CAPTURE,
            CaptureMethod.APPLICATION_FRAME_HOST
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
                elif method == CaptureMethod.WGC_API:
                    img = self._capture_wgc_api(hwnd)
                elif method == CaptureMethod.DWM_API:
                    img = self._capture_dwm_api(hwnd)
                elif method == CaptureMethod.GDIPLUS:
                    img = self._capture_gdiplus(hwnd)
                elif method == CaptureMethod.DIRECTX_CAPTURE:
                    img = self._capture_directx(hwnd)
                elif method == CaptureMethod.WINDOWS_SCREENSHOT:
                    img = self._capture_windows_screenshot(hwnd)
                elif method == CaptureMethod.DXGI_SURFACE_SHARING:
                    img = self._capture_dxgi_surface_sharing(hwnd)
                elif method == CaptureMethod.WINDOWS_MEDIA_CAPTURE:
                    img = self._capture_windows_media(hwnd)
                elif method == CaptureMethod.D3D9_SURFACE_CAPTURE:
                    img = self._capture_d3d9_surface(hwnd)
                elif method == CaptureMethod.APPLICATION_FRAME_HOST:
                    img = self._capture_application_frame_host(hwnd)
                
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
    
    def _capture_wgc_api(self, hwnd: int) -> Optional[np.ndarray]:
        """方法11: 使用Windows.Graphics.Capture API (Win10 1809+)"""
        try:
            import winrt.windows.graphics.capture as wgc
            import winrt.windows.graphics.directx as wgdx
            import winrt.windows.graphics.directx.direct3d11 as wgdd
            
            # 初始化WinRT
            pythoncom.CoInitialize()
            
            # 创建捕获项
            interop = wgc.GraphicsCaptureItem
            item = interop.create_from_window(hwnd)
            
            if not item:
                raise Exception("Failed to create capture item")
            
            # 创建帧池和会话
            device = wgdx.Direct3D11.create_device()
            frame_pool = wgc.Direct3D11CaptureFramePool.create_free_threaded(
                device,
                wgdx.DirectXPixelFormat.B8G8R8A8_UINT,
                1,
                item.size
            )
            
            session = frame_pool.create_capture_session(item)
            session.start_capture()
            
            # 等待一帧
            frame = None
            for _ in range(10):  # 最多等待10次
                frame = frame_pool.try_get_next_frame()
                if frame:
                    break
                time.sleep(0.1)
            
            if not frame:
                raise Exception("Failed to get frame")
            
            # 获取帧内容
            surface = wgdd.IDirect3DSurface
            texture = surface.get_texture()
            
            # 转换为numpy数组
            with texture.map() as mapping:
                img = np.frombuffer(
                    mapping.data,
                    dtype=np.uint8,
                    count=mapping.data_length
                ).reshape((item.size.height, item.size.width, 4))
            
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
            
            # 清理资源
            session.close()
            frame_pool.close()
            pythoncom.CoUninitialize()
            
            return img
            
        except Exception as e:
            self.logger.debug(f"Windows.Graphics.Capture API method failed: {e}")
            return None
    
    def _capture_dwm_api(self, hwnd: int) -> Optional[np.ndarray]:
        """方法12: 使用DWM API捕获窗口"""
        try:
            # 获取窗口矩形
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            
            # 创建兼容DC和位图
            hdc = win32gui.GetWindowDC(0)
            mfcDC = win32ui.CreateDCFromHandle(hdc)
            saveDC = mfcDC.CreateCompatibleDC()
            
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            
            saveDC.SelectObject(saveBitMap)
            
            # 使用DWM API捕获窗口
            dwm = ctypes.windll.dwmapi
            dwm.DwmGetWindowAttribute(
                hwnd,
                9,  # DWMWA_EXTENDED_FRAME_BOUNDS
                ctypes.byref(rect),
                ctypes.sizeof(rect)
            )
            
            # 捕获到内存位图
            result = saveDC.BitBlt(
                (0, 0),
                (width, height),
                mfcDC,
                (rect[0], rect[1]),
                win32con.SRCCOPY
            )
            
            if result:
                bmpinfo = saveBitMap.GetInfo()
                bmpstr = saveBitMap.GetBitmapBits(True)
                img = np.frombuffer(bmpstr, dtype='uint8').reshape((bmpinfo['bmHeight'], bmpinfo['bmWidth'], 4))
                img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
                
                win32gui.DeleteObject(saveBitMap.GetHandle())
                saveDC.DeleteDC()
                mfcDC.DeleteDC()
                win32gui.ReleaseDC(0, hdc)
                
                return img
            return None
            
        except Exception as e:
            self.logger.debug(f"DWM API method failed: {e}")
            return None
    
    def _capture_gdiplus(self, hwnd: int) -> Optional[np.ndarray]:
        """方法13: 使用GDI+捕获窗口"""
        try:
            import clr
            clr.AddReference("System.Drawing")
            
            from System.Drawing import Graphics, Bitmap, Rectangle
            from System.Windows.Forms import Control
            
            # 获取窗口矩形
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            
            # 创建GDI+位图
            bmp = Bitmap(width, height)
            graphics = Graphics.FromImage(bmp)
            
            # 获取窗口DC
            hdc = graphics.GetHdc()
            
            try:
                # 捕获窗口内容
                success = ctypes.windll.user32.PrintWindow(
                    hwnd,
                    hdc,
                    2  # PW_CLIENTONLY
                )
                
                if not success:
                    raise Exception("PrintWindow failed")
                
                # 转换为numpy数组
                img = np.zeros((height, width, 3), dtype=np.uint8)
                
                for y in range(height):
                    for x in range(width):
                        color = bmp.GetPixel(x, y)
                        img[y, x] = [color.B, color.G, color.R]
                
                return img
                
            finally:
                graphics.ReleaseHdc(hdc)
                graphics.Dispose()
                bmp.Dispose()
                
        except Exception as e:
            self.logger.debug(f"GDI+ method failed: {e}")
            return None
    
    def _capture_directx(self, hwnd: int) -> Optional[np.ndarray]:
        """方法14: 使用DirectX捕获窗口"""
        try:
            # 初始化Direct3D9
            d3d = ctypes.windll.d3d9
            if not d3d:
                raise Exception("Failed to load d3d9.dll")
            
            # 创建Direct3D设备
            d3d_obj = d3d.Direct3DCreate9(32)  # D3D_SDK_VERSION
            if not d3d_obj:
                raise Exception("Direct3DCreate9 failed")
            
            # 获取窗口矩形
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            
            # 创建设备参数
            present_params = D3DPRESENT_PARAMETERS()
            present_params.Windowed = True
            present_params.SwapEffect = 1  # D3DSWAPEFFECT_DISCARD
            present_params.BackBufferFormat = 0  # D3DFMT_UNKNOWN
            present_params.BackBufferWidth = width
            present_params.BackBufferHeight = height
            
            # 创建设备
            device = ctypes.POINTER(IDirect3DDevice9)()
            hr = d3d_obj.CreateDevice(
                0,  # D3DADAPTER_DEFAULT
                1,  # D3DDEVTYPE_HAL
                hwnd,
                0x00000020,  # D3DCREATE_SOFTWARE_VERTEXPROCESSING
                ctypes.byref(present_params),
                ctypes.byref(device)
            )
            
            if hr != 0:
                raise Exception(f"CreateDevice failed with HRESULT: {hr}")
            
            # 创建表面
            surface = ctypes.POINTER(IDirect3DSurface9)()
            hr = device.CreateOffscreenPlainSurface(
                width,
                height,
                22,  # D3DFMT_A8R8G8B8
                0,   # D3DPOOL_DEFAULT
                ctypes.byref(surface),
                None
            )
            
            if hr != 0:
                raise Exception(f"CreateOffscreenPlainSurface failed with HRESULT: {hr}")
            
            # 捕获到表面
            hr = device.GetFrontBufferData(0, surface)
            if hr != 0:
                raise Exception(f"GetFrontBufferData failed with HRESULT: {hr}")
            
            # 锁定表面
            locked_rect = D3DLOCKED_RECT()
            hr = surface.LockRect(ctypes.byref(locked_rect), None, 0)
            if hr != 0:
                raise Exception(f"LockRect failed with HRESULT: {hr}")
            
            # 转换为numpy数组
            img = np.frombuffer(
                ctypes.string_at(locked_rect.pBits, width * height * 4),
                dtype=np.uint8
            ).reshape((height, width, 4))
            
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
            
            # 解锁表面
            surface.UnlockRect()
            
            # 释放资源
            surface.Release()
            device.Release()
            d3d_obj.Release()
            
            return img
            
        except Exception as e:
            self.logger.debug(f"DirectX method failed: {e}")
            return None
    
    def _capture_windows_screenshot(self, hwnd: int) -> Optional[np.ndarray]:
        """方法15: 使用Windows Screenshot API"""
        try:
            # 获取窗口矩形
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            
            # 使用Windows.Graphics.Capture API
            import winrt.windows.graphics.capture as wgc
            import winrt.windows.graphics.directx as wgdx
            import winrt.windows.graphics.directx.direct3d11 as wgdd
            
            # 初始化WinRT
            pythoncom.CoInitialize()
            
            # 创建捕获项
            interop = wgc.GraphicsCaptureItem
            item = interop.create_from_window(hwnd)
            
            if not item:
                raise Exception("Failed to create capture item")
            
            # 创建帧池和会话
            device = wgdx.Direct3D11.create_device()
            frame_pool = wgc.Direct3D11CaptureFramePool.create_free_threaded(
                device,
                wgdx.DirectXPixelFormat.B8G8R8A8_UINT,
                1,
                item.size
            )
            
            session = frame_pool.create_capture_session(item)
            session.start_capture()
            
            # 等待一帧
            frame = None
            for _ in range(10):  # 最多等待10次
                frame = frame_pool.try_get_next_frame()
                if frame:
                    break
                time.sleep(0.1)
            
            if not frame:
                raise Exception("Failed to get frame")
            
            # 获取帧内容
            surface = wgdd.IDirect3DSurface
            texture = surface.get_texture()
            
            # 转换为numpy数组
            with texture.map() as mapping:
                img = np.frombuffer(
                    mapping.data,
                    dtype=np.uint8,
                    count=mapping.data_length
                ).reshape((item.size.height, item.size.width, 4))
            
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
            
            # 清理资源
            session.close()
            frame_pool.close()
            pythoncom.CoUninitialize()
            
            return img
            
        except Exception as e:
            self.logger.debug(f"Windows Screenshot API method failed: {e}")
            return None
    
    def _capture_dxgi_surface_sharing(self, hwnd: int) -> Optional[np.ndarray]:
        """方法16: 使用DXGI表面共享捕获"""
        try:
            import dxgi
            
            # 获取窗口矩形
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            
            # 创建DXGI工厂
            factory = dxgi.create_dxgi_factory()
            
            # 获取适配器
            adapter = factory.enum_adapters(0)
            
            # 创建设备
            device = dxgi.create_device(adapter)
            
            # 创建输出
            output = adapter.enum_outputs(0)
            
            # 创建共享表面
            surface = device.create_shared_surface(
                width,
                height,
                dxgi.FORMAT_B8G8R8A8_UNORM
            )
            
            # 捕获到表面
            output.duplicate_output(device, surface)
            
            # 映射表面
            mapped_rect = surface.map()
            
            # 转换为numpy数组
            img = np.frombuffer(
                mapped_rect.pBits,
                dtype=np.uint8,
                count=width * height * 4
            ).reshape((height, width, 4))
            
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
            
            # 取消映射
            surface.unmap()
            
            # 释放资源
            surface.release()
            output.release()
            adapter.release()
            factory.release()
            
            return img
            
        except Exception as e:
            self.logger.debug(f"DXGI Surface Sharing method failed: {e}")
            return None
    
    def _capture_windows_media(self, hwnd: int) -> Optional[np.ndarray]:
        """方法17: 使用Windows.Media.Capture API"""
        try:
            import winrt.windows.media.capture as wmc
            import winrt.windows.graphics.imaging as wgi
            import winrt.windows.storage.streams as wss
            
            # 初始化WinRT
            pythoncom.CoInitialize()
            
            # 创建捕获对象
            capture = wmc.MediaCapture()
            
            # 初始化设置
            settings = wmc.MediaCaptureInitializationSettings()
            settings.video_device_id = ""  # 默认设备
            settings.streaming_capture_mode = wmc.StreamingCaptureMode.VIDEO
            
            # 异步初始化
            import asyncio
            async def init_capture():
                await capture.initialize_async(settings)
                
            asyncio.run(init_capture())
            
            # 创建内存流
            stream = wss.InMemoryRandomAccessStream()
            
            # 捕获到流
            async def capture_to_stream():
                await capture.capture_photo_to_stream_async(
                    wgi.BitmapEncoder.png_encoder_id(),
                    stream
                )
                
            asyncio.run(capture_to_stream())
            
            # 从流中读取数据
            stream.seek(0)
            reader = wss.DataReader(stream.get_input_stream_at(0))
            reader.load_async(stream.size)
            
            # 转换为numpy数组
            buffer = reader.read_buffer(stream.size)
            img_bytes = bytearray(buffer)
            img = cv2.imdecode(np.frombuffer(img_bytes, np.uint8), cv2.IMREAD_COLOR)
            
            # 释放资源
            stream.close()
            capture.close()
            pythoncom.CoUninitialize()
            
            return img
            
        except Exception as e:
            self.logger.debug(f"Windows.Media.Capture API method failed: {e}")
            return None
    
    def _capture_d3d9_surface(self, hwnd: int) -> Optional[np.ndarray]:
        """方法18: 使用Direct3D9表面捕获"""
        try:
            if not self.d3d9:
                raise Exception("D3D9 not available")
            
            # 获取窗口矩形
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            
            # 创建Direct3D9对象
            d3d = self.d3d9.Direct3DCreate9(32)  # D3D_SDK_VERSION
            if not d3d:
                raise Exception("Direct3DCreate9 failed")
            
            # 创建设备参数
            present_params = D3DPRESENT_PARAMETERS()
            present_params.Windowed = True
            present_params.SwapEffect = 1  # D3DSWAPEFFECT_DISCARD
            present_params.BackBufferFormat = 0  # D3DFMT_UNKNOWN
            present_params.BackBufferWidth = width
            present_params.BackBufferHeight = height
            
            # 创建设备
            device = ctypes.POINTER(IDirect3DDevice9)()
            hr = d3d.CreateDevice(
                0,  # D3DADAPTER_DEFAULT
                1,  # D3DDEVTYPE_HAL
                hwnd,
                0x00000020,  # D3DCREATE_SOFTWARE_VERTEXPROCESSING
                ctypes.byref(present_params),
                ctypes.byref(device)
            )
            
            if hr != 0:
                raise Exception(f"CreateDevice failed with HRESULT: {hr}")
            
            # 创建表面
            surface = ctypes.POINTER(IDirect3DSurface9)()
            hr = device.CreateOffscreenPlainSurface(
                width,
                height,
                22,  # D3DFMT_A8R8G8B8
                0,   # D3DPOOL_DEFAULT
                ctypes.byref(surface),
                None
            )
            
            if hr != 0:
                raise Exception(f"CreateOffscreenPlainSurface failed with HRESULT: {hr}")
            
            # 捕获到表面
            hr = device.GetFrontBufferData(0, surface)
            if hr != 0:
                raise Exception(f"GetFrontBufferData failed with HRESULT: {hr}")
            
            # 锁定表面
            locked_rect = D3DLOCKED_RECT()
            hr = surface.LockRect(ctypes.byref(locked_rect), None, 0)
            if hr != 0:
                raise Exception(f"LockRect failed with HRESULT: {hr}")
            
            # 转换为numpy数组
            img = np.frombuffer(
                ctypes.string_at(locked_rect.pBits, width * height * 4),
                dtype=np.uint8
            ).reshape((height, width, 4))
            
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
            
            # 解锁表面
            surface.UnlockRect()
            
            # 释放资源
            surface.Release()
            device.Release()
            d3d.Release()
            
            return img
            
        except Exception as e:
            self.logger.debug(f"Direct3D9 Surface Capture method failed: {e}")
            return None
    
    def _capture_windows_ai_vision(self, hwnd: int) -> Optional[np.ndarray]:
        """方法19: 使用Windows AI Vision API"""
        try:
            import winrt.windows.ai.machinelearning as waml
            import winrt.windows.media as wm
            import winrt.windows.graphics.imaging as wgi
            
            # 首先使用其他方法捕获窗口
            img = self._capture_bitblt(hwnd)
            if img is None:
                raise Exception("Failed to capture window")
            
            # 转换为Windows.Media.VideoFrame
            height, width, _ = img.shape
            
            # 创建软件位图
            software_bitmap = wgi.SoftwareBitmap.create(
                wgi.BitmapPixelFormat.B8G8R8A8,
                width,
                height
            )
            
            # 复制数据
            buffer = software_bitmap.lock_buffer(wgi.BitmapBufferAccessMode.READ_WRITE)
            plane = buffer.get_plane(0)
            
            import ctypes
            ctypes.memmove(
                plane.buffer,
                img.ctypes.data,
                width * height * 4
            )
            
            buffer.close()
            
            # 创建视频帧
            video_frame = wm.VideoFrame.create_with_software_bitmap(software_bitmap)
            
            # 这里可以添加AI处理逻辑
            # 例如使用Windows.AI.MachineLearning进行图像分析
            
            # 转换回numpy数组
            result_img = np.zeros((height, width, 3), dtype=np.uint8)
            
            # 从视频帧获取数据
            with video_frame.as_buffer() as buffer:
                data = buffer.to_array()
                result_img = np.frombuffer(data, dtype=np.uint8).reshape((height, width, 4))
                result_img = cv2.cvtColor(result_img, cv2.COLOR_BGRA2BGR)
            
            return result_img
            
        except Exception as e:
            self.logger.debug(f"Windows AI Vision API method failed: {e}")
            return None
    
    def _capture_application_frame_host(self, hwnd: int) -> Optional[np.ndarray]:
        """方法20: 捕获ApplicationFrameHost窗口"""
        try:
            # 检查是否是UWP窗口
            class_name = win32gui.GetClassName(hwnd)
            if class_name != "ApplicationFrameWindow":
                raise Exception("Not an ApplicationFrameWindow")
            
            # 查找子窗口
            def enum_child_windows(hwnd, lparam):
                class_name = win32gui.GetClassName(hwnd)
                if class_name == "Windows.UI.Core.CoreWindow":
                    lparam.append(hwnd)
                return True
            
            child_windows = []
            win32gui.EnumChildWindows(hwnd, enum_child_windows, child_windows)
            
            if not child_windows:
                raise Exception("No CoreWindow found")
            
            # 使用BitBlt捕获子窗口
            return self._capture_bitblt(child_windows[0])
            
        except Exception as e:
            self.logger.debug(f"ApplicationFrameHost method failed: {e}")
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