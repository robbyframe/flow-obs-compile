"""
Sistem Automasi Integrasi PowerPoint/Keynote dan OBS (Cross-Platform)
Mengubah scene OBS berdasarkan notes di slide PowerPoint/Keynote
Support: Windows (PowerPoint), macOS (PowerPoint & Keynote)

VERSION: 1.0 FINAL
"""

import time
import sys
import re
import platform
import subprocess
from typing import Optional, List
from abc import ABC, abstractmethod

# Detect platform
IS_WINDOWS = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"

# Import platform-specific modules
if IS_WINDOWS:
    try:
        import win32com.client
        from win32com.client import constants
    except ImportError:
        print("Error: pywin32 tidak terinstall. Install dengan: pip install pywin32")
        sys.exit(1)

try:
    import obsws_python as obs
except ImportError:
    print("Error: obsws-python tidak terinstall. Install dengan: pip install obsws-python")
    sys.exit(1)


# ============================================================================
# ABSTRACT BASE CLASS - PRESENTATION CONTROLLER
# ============================================================================

class PresentationController(ABC):
    """Abstract base class untuk presentation controller"""
    
    def __init__(self, debug: bool = False):
        self.debug = debug
        self.current_slide_index = None
    
    @abstractmethod
    def connect(self) -> bool:
        """Koneksi ke aplikasi presentasi"""
        pass
    
    @abstractmethod
    def get_current_slide_index(self) -> Optional[int]:
        """Mendapatkan index slide yang sedang aktif"""
        pass
    
    @abstractmethod
    def get_slide_notes(self, slide_index: int) -> Optional[str]:
        """Mendapatkan notes dari slide tertentu"""
        pass
    
    @abstractmethod
    def get_presentation_name(self) -> str:
        """Mendapatkan nama presentasi yang aktif"""
        pass
    
    @abstractmethod
    def is_in_slideshow(self) -> bool:
        """Cek apakah sedang dalam mode slideshow"""
        pass


# ============================================================================
# WINDOWS - POWERPOINT CONTROLLER
# ============================================================================

class WindowsPowerPointController(PresentationController):
    """Controller untuk PowerPoint di Windows"""
    
    def __init__(self, debug: bool = False):
        super().__init__(debug)
        self.ppt_app = None
        self.is_slideshow_active = False
    
    def connect(self) -> bool:
        try:
            self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            self.ppt_app.Visible = True
            if self.ppt_app.Presentations.Count == 0:
                print("✗ Tidak ada presentasi PowerPoint yang terbuka")
                return False
            presentation = self.ppt_app.ActivePresentation
            print(f"✓ Terhubung ke PowerPoint: {presentation.Name}")
            return True
        except Exception as e:
            print(f"✗ Error koneksi ke PowerPoint: {e}")
            return False
    
    def is_in_slideshow(self) -> bool:
        try:
            if not self.ppt_app or self.ppt_app.Presentations.Count == 0:
                return False
            if hasattr(self.ppt_app, 'SlideShowWindows') and self.ppt_app.SlideShowWindows.Count > 0:
                self.is_slideshow_active = True
                return True
            self.is_slideshow_active = False
            return False
        except:
            self.is_slideshow_active = False
            return False
    
    def get_current_slide_index(self) -> Optional[int]:
        try:
            if not self.ppt_app or self.ppt_app.Presentations.Count == 0:
                return None
            if not self.is_in_slideshow():
                return None
            if hasattr(self.ppt_app, 'SlideShowWindows') and self.ppt_app.SlideShowWindows.Count > 0:
                slideshow = self.ppt_app.SlideShowWindows(1)
                for prop in ['View.Slide', 'View.CurrentSlide', 'CurrentSlide']:
                    try:
                        if prop == 'View.Slide':
                            slide = slideshow.View.Slide
                        elif prop == 'View.CurrentSlide':
                            slide = slideshow.View.CurrentSlide
                        else:
                            slide = slideshow.CurrentSlide
                        if slide:
                            return slide.SlideIndex
                    except:
                        continue
            return None
        except Exception as e:
            if self.debug:
                print(f"[Debug] Error get_current_slide_index: {e}")
            return None
    
    def get_slide_notes(self, slide_index: int) -> Optional[str]:
        try:
            if not self.ppt_app or self.ppt_app.Presentations.Count == 0:
                return None
            presentation = self.ppt_app.ActivePresentation
            if slide_index < 1 or slide_index > presentation.Slides.Count:
                return None
            slide = presentation.Slides.Item(slide_index)
            notes_text = ""
            if slide.HasNotesPage:
                notes_page = slide.NotesPage
                for shape in notes_page.Shapes:
                    if shape.HasTextFrame and shape.TextFrame.HasText:
                        text = shape.TextFrame.TextRange.Text
                        if text:
                            notes_text += text + "\n"
            return notes_text.strip() if notes_text else None
        except Exception as e:
            if self.debug:
                print(f"[Debug] Error membaca notes: {e}")
            return None
    
    def get_presentation_name(self) -> str:
        try:
            if self.ppt_app and self.ppt_app.Presentations.Count > 0:
                return self.ppt_app.ActivePresentation.Name
        except:
            pass
        return "Unknown"


# ============================================================================
# MACOS - POWERPOINT CONTROLLER
# ============================================================================

class MacPowerPointController(PresentationController):
    """Controller untuk PowerPoint di macOS menggunakan AppleScript"""
    
    def __init__(self, debug: bool = False):
        super().__init__(debug)
        self.is_connected = False
        self.is_slideshow_active = False
    
    def _run_applescript(self, script: str) -> Optional[str]:
        try:
            result = subprocess.run(
                ['osascript', '-e', script],
                capture_output=True,
                text=True,
                timeout=2
            )
            if result.returncode == 0:
                return result.stdout.strip()
            else:
                if self.debug:
                    print(f"[Debug] AppleScript error: {result.stderr.strip()}")
                return None
        except subprocess.TimeoutExpired:
            if self.debug:
                print(f"[Debug] AppleScript timeout")
            return None
        except Exception as e:
            if self.debug:
                print(f"[Debug] Error running AppleScript: {e}")
            return None
    
    def connect(self) -> bool:
        script = '''
        tell application "Microsoft PowerPoint"
            if it is running then
                try
                    set presCount to count of presentations
                    if presCount > 0 then
                        set presName to name of active presentation
                        return presName
                    else
                        return "NO_PRESENTATION"
                    end if
                on error
                    return "ERROR"
                end try
            else
                return "NOT_RUNNING"
            end if
        end tell
        '''
        result = self._run_applescript(script)
        if result == "NOT_RUNNING":
            print("✗ PowerPoint tidak berjalan")
            return False
        elif result == "NO_PRESENTATION":
            print("✗ Tidak ada presentasi PowerPoint yang terbuka")
            return False
        elif result == "ERROR" or not result:
            print("✗ Error koneksi ke PowerPoint")
            return False
        else:
            print(f"✓ Terhubung ke PowerPoint: {result}")
            print("ℹ Tips: Tekan ⌥⌘P (Option+Command+P) untuk slideshow")
            self.is_connected = True
            return True
    
    def is_in_slideshow(self) -> bool:
        if not self.is_connected:
            return False
        for script in [
            'tell application "Microsoft PowerPoint"\ntry\nif (count of slide show windows) > 0 then\nreturn "true"\nelse\nreturn "false"\nend if\non error\nreturn "false"\nend try\nend tell',
            'tell application "Microsoft PowerPoint"\ntry\nset slideshowView to slide show view of slide show window 1\nif slideshowView exists then\nreturn "true"\nelse\nreturn "false"\nend if\non error\nreturn "false"\nend try\nend tell',
            'tell application "Microsoft PowerPoint"\ntry\nset currentSlide to slide index of slide of slide show view of slide show window 1\nreturn "true"\non error\nreturn "false"\nend try\nend tell'
        ]:
            result = self._run_applescript(script)
            if result == "true":
                self.is_slideshow_active = True
                return True
        self.is_slideshow_active = False
        return False
    
    def get_current_slide_index(self) -> Optional[int]:
        if not self.is_connected:
            return None
        if not self.is_in_slideshow():
            return None
        script = '''
        tell application "Microsoft PowerPoint"
            try
                set currentSlide to slide index of slide of slide show view of slide show window 1
                return currentSlide as text
            on error
                return "ERROR"
            end try
        end tell
        '''
        result = self._run_applescript(script)
        if result and result != "ERROR":
            try:
                return int(result)
            except:
                pass
        return None
    
    def get_slide_notes(self, slide_index: int) -> Optional[str]:
        if not self.is_connected:
            return None
        for script in [
            f'tell application "Microsoft PowerPoint"\ntry\ntell active presentation\nset theSlide to slide {slide_index}\nset notesText to content of text range of text frame of shape 2 of notes page of theSlide\nreturn notesText\nend tell\non error\nreturn "ERROR1"\nend try\nend tell',
            f'tell application "Microsoft PowerPoint"\ntry\ntell active presentation\nset theSlide to slide {slide_index}\nset notesPage to notes page of theSlide\nset notesText to ""\nrepeat with shp in shapes of notesPage\ntry\nif has text frame of shp then\nif has text of text frame of shp then\nset notesText to notesText & (content of text range of text frame of shp) & " "\nend if\nend if\nend try\nend repeat\nreturn notesText\nend tell\non error\nreturn ""\nend try\nend tell',
            f'tell application "Microsoft PowerPoint"\ntry\nset theSlide to slide {slide_index} of active presentation\nset notesText to content of text range of text frame of shape 2 of notes page of theSlide\nreturn notesText\non error\nreturn ""\nend try\nend tell'
        ]:
            result = self._run_applescript(script)
            if result and result not in ["ERROR1", ""]:
                return result
        return None
    
    def get_presentation_name(self) -> str:
        script = '''
        tell application "Microsoft PowerPoint"
            try
                return name of active presentation
            end try
        end tell
        '''
        result = self._run_applescript(script)
        return result if result else "Unknown"


# ============================================================================
# MACOS - KEYNOTE CONTROLLER
# ============================================================================

class KeynoteController(PresentationController):
    """Controller untuk Keynote di macOS menggunakan AppleScript"""
    
    def __init__(self, debug: bool = False):
        super().__init__(debug)
        self.is_connected = False
        self.is_slideshow_active = False
    
    def _run_applescript(self, script: str) -> Optional[str]:
        try:
            result = subprocess.run(
                ['osascript', '-e', script],
                capture_output=True,
                text=True,
                timeout=2
            )
            if result.returncode == 0:
                return result.stdout.strip()
            else:
                if self.debug:
                    print(f"[Debug] AppleScript error: {result.stderr}")
                return None
        except Exception as e:
            if self.debug:
                print(f"[Debug] Error running AppleScript: {e}")
            return None
    
    def connect(self) -> bool:
        script = '''
        tell application "Keynote"
            if it is running then
                try
                    if (count of documents) > 0 then
                        set docName to name of front document
                        return docName
                    else
                        return "NO_DOCUMENT"
                    end if
                on error
                    return "ERROR"
                end try
            else
                return "NOT_RUNNING"
            end if
        end tell
        '''
        result = self._run_applescript(script)
        if result == "NOT_RUNNING":
            print("✗ Keynote tidak berjalan")
            return False
        elif result == "NO_DOCUMENT":
            print("✗ Tidak ada presentasi Keynote yang terbuka")
            return False
        elif result == "ERROR" or not result:
            print("✗ Error koneksi ke Keynote")
            return False
        else:
            print(f"✓ Terhubung ke Keynote: {result}")
            print("ℹ Tips: Tekan ⌥⌘P (Option+Command+P) untuk play slideshow")
            self.is_connected = True
            return True
    
    def is_in_slideshow(self) -> bool:
        if not self.is_connected:
            return False
        script = '''
        tell application "Keynote"
            try
                if playing is true then
                    return "true"
                else
                    return "false"
                end if
            on error
                return "false"
            end try
        end tell
        '''
        result = self._run_applescript(script)
        self.is_slideshow_active = (result == "true")
        return self.is_slideshow_active
    
    def get_current_slide_index(self) -> Optional[int]:
        if not self.is_connected:
            return None
        if not self.is_in_slideshow():
            return None
        script = '''
        tell application "Keynote"
            try
                set currentSlideNum to slide number of current slide of front document
                return currentSlideNum as text
            on error
                return "ERROR"
            end try
        end tell
        '''
        result = self._run_applescript(script)
        if result and result != "ERROR":
            try:
                return int(result)
            except:
                pass
        return None
    
    def get_slide_notes(self, slide_index: int) -> Optional[str]:
        if not self.is_connected:
            return None
        script = f'''
        tell application "Keynote"
            try
                set theSlide to slide {slide_index} of front document
                set notesText to presenter notes of theSlide
                return notesText
            on error
                return ""
            end try
        end tell
        '''
        result = self._run_applescript(script)
        return result if result else None
    
    def get_presentation_name(self) -> str:
        script = '''
        tell application "Keynote"
            try
                return name of front document
            end try
        end tell
        '''
        result = self._run_applescript(script)
        return result if result else "Unknown"


# ============================================================================
# OBS INTEGRATION CLASS
# ============================================================================

class PowerPointOBSIntegrator:
    """Main integrator class untuk PowerPoint/Keynote dan OBS"""
    
    def __init__(self, 
                 presentation_type: str = "auto",
                 obs_host: str = "localhost", 
                 obs_port: int = 4455, 
                 obs_password: str = "", 
                 debug: bool = False):
        self.obs_host = obs_host
        self.obs_port = obs_port
        self.obs_password = obs_password
        self.debug = debug
        self.obs_client = None
        self.available_scenes = []
        self.controller = self._init_controller(presentation_type)
    
    def _init_controller(self, presentation_type: str) -> Optional[PresentationController]:
        if IS_WINDOWS:
            if presentation_type in ["auto", "powerpoint"]:
                print("Platform: Windows - Menggunakan PowerPoint")
                return WindowsPowerPointController(debug=self.debug)
            else:
                print("✗ Keynote tidak tersedia di Windows")
                return None
        elif IS_MAC:
            if presentation_type == "keynote":
                print("Platform: macOS - Menggunakan Keynote")
                return KeynoteController(debug=self.debug)
            elif presentation_type == "powerpoint":
                print("Platform: macOS - Menggunakan PowerPoint")
                return MacPowerPointController(debug=self.debug)
            else:
                print("Platform: macOS - Auto-detect...")
                keynote = KeynoteController(debug=self.debug)
                if keynote.connect():
                    return keynote
                ppt = MacPowerPointController(debug=self.debug)
                if ppt.connect():
                    return ppt
                print("✗ Tidak ada Keynote atau PowerPoint yang aktif")
                return None
        else:
            print(f"✗ Platform tidak didukung: {platform.system()}")
            return None
    
    def connect_obs(self) -> bool:
        try:
            self.obs_client = obs.ReqClient(
                host=self.obs_host,
                port=self.obs_port,
                password=self.obs_password
            )
            scenes = self.obs_client.get_scene_list()
            self.available_scenes = [scene['sceneName'] for scene in scenes.scenes]
            print(f"✓ Terhubung ke OBS. Scene tersedia: {', '.join(self.available_scenes)}")
            return True
        except Exception as e:
            print(f"✗ Error koneksi ke OBS: {e}")
            print("Pastikan OBS Studio terbuka dan WebSocket Server diaktifkan:")
            print("  Tools > WebSocket Server Settings > Enable WebSocket server")
            return False
    
    def extract_scene_name(self, notes: str) -> Optional[str]:
        if not notes or not self.available_scenes:
            return None
        notes_clean = re.sub(r'\s+', ' ', notes.strip())
        notes_lower = notes_clean.lower().strip()
        for scene_name in self.available_scenes:
            if scene_name.lower().strip() == notes_lower:
                return scene_name
        scenes_sorted = sorted(self.available_scenes, key=len, reverse=True)
        for scene_name in scenes_sorted:
            pattern = r'\b' + re.escape(scene_name) + r'\b'
            if re.search(pattern, notes_clean, re.IGNORECASE):
                return scene_name
        for scene_name in scenes_sorted:
            scene_lower = scene_name.lower().strip()
            if scene_lower in notes_lower:
                pattern = r'(?:^|\s|[^\w])' + re.escape(scene_name) + r'(?:\s|[^\w]|$)'
                if re.search(pattern, notes_clean, re.IGNORECASE):
                    return scene_name
        return None
    
    def switch_obs_scene(self, scene_name: str) -> bool:
        try:
            self.obs_client.set_current_program_scene(sceneName=scene_name)
            print(f"✓ Scene diubah ke: {scene_name}")
            return True
        except:
            try:
                self.obs_client.set_current_program_scene(scene_name)
                print(f"✓ Scene diubah ke: {scene_name}")
                return True
            except Exception as e:
                print(f"✗ Error mengubah scene: {e}")
                return False
    
    def monitor_and_sync(self):
        print("\n" + "="*60)
        print("Sistem Monitoring PowerPoint/Keynote-OBS Aktif")
        print("Tekan Ctrl+C untuk menghentikan")
        print("="*60 + "\n")
        
        if not self.controller:
            print("✗ Controller tidak tersedia")
            return
        
        if IS_WINDOWS and not self.controller.is_connected if hasattr(self.controller, 'is_connected') else True:
            if not self.controller.connect():
                return
        
        if not self.connect_obs():
            return
        
        print("\n⏳ Menunggu slideshow dimulai...")
        if IS_MAC:
            print("   Tips: Tekan ⌥⌘P (Option+Command+P) untuk memulai slideshow")
        else:
            print("   Tips: Tekan F5 untuk memulai slideshow")
        
        if self.debug:
            print("   [Debug mode aktif - akan menampilkan detail teknis]\n")
        else:
            print()
        
        last_slide_index = None
        last_slideshow_state = False
        check_count = 0
        slideshow_warned = False
        
        try:
            while True:
                check_count += 1
                is_slideshow = self.controller.is_in_slideshow()
                
                if self.debug and check_count % 4 == 0:
                    print(f"[Debug] Check #{check_count}: Slideshow={is_slideshow}")
                
                if is_slideshow != last_slideshow_state:
                    if is_slideshow:
                        print("✓ Slideshow dimulai - Monitoring aktif\n")
                        slideshow_warned = False
                        last_slide_index = None
                    else:
                        print("\n⏸ Slideshow berhenti - Monitoring dijeda")
                        print("  (OBS scene tidak akan berubah saat edit mode)\n")
                        last_slide_index = None
                    last_slideshow_state = is_slideshow
                
                if is_slideshow:
                    new_index = self.controller.get_current_slide_index()
                    if new_index and new_index != last_slide_index:
                        last_slide_index = new_index
                        print(f"→ Slide {new_index}")
                        notes = self.controller.get_slide_notes(new_index)
                        if notes:
                            notes_preview = notes[:80] + "..." if len(notes) > 80 else notes
                            print(f"  Notes: {notes_preview}")
                            scene_name = self.extract_scene_name(notes)
                            if scene_name:
                                self.switch_obs_scene(scene_name)
                            else:
                                print(f"  ⚠ Scene tidak ditemukan")
                        else:
                            print(f"  ⚠ Tidak ada notes")
                else:
                    if not slideshow_warned and check_count > 10:
                        slideshow_warned = True
                
                time.sleep(0.5)
                
        except KeyboardInterrupt:
            print("\n\n✓ Monitoring dihentikan")
        except Exception as e:
            print(f"\n✗ Error: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def main():
    print("="*60)
    print("Sistem Automasi Integrasi PowerPoint/Keynote dan OBS")
    print("="*60)
    print(f"\nPlatform: {platform.system()}")
    print("\nKonfigurasi:")
    print("  - OBS WebSocket harus diaktifkan")
    print("  - PowerPoint/Keynote harus terbuka dengan presentasi aktif")
    print("  - Notes/Presenter Notes di slide harus berisi nama scene OBS")
    print()
    
    obs_host = "localhost"
    obs_port = 4455
    obs_password = ""
    debug_mode = False
    presentation_type = "auto"
    
    integrator = PowerPointOBSIntegrator(
        presentation_type=presentation_type,
        obs_host=obs_host,
        obs_port=obs_port,
        obs_password=obs_password,
        debug=debug_mode
    )
    
    integrator.monitor_and_sync()


if __name__ == "__main__":
    main()
