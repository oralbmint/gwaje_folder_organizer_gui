import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import queue

class FolderOrganizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("과제 폴더 정리 도구")
        self.root.geometry("600x450")
        self.root.resizable(True, True)
        
        self.selected_folder = tk.StringVar()
        self.log_queue = queue.Queue()
        self.is_running = False
        
        self.create_widgets()
        self.update_log()
        
    def create_widgets(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 폴더 선택 프레임
        folder_frame = ttk.LabelFrame(main_frame, text="폴더 선택", padding="10")
        folder_frame.pack(fill=tk.X, pady=5)
        
        # 폴더 경로 표시 및 선택 버튼
        ttk.Entry(folder_frame, textvariable=self.selected_folder, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(folder_frame, text="폴더 찾기", command=self.browse_folder).pack(side=tk.RIGHT, padx=5)
        
        # 로그 표시 영역
        log_frame = ttk.LabelFrame(main_frame, text="작업 로그", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(log_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 로그 텍스트 위젯
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)
        
        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        # 실행 버튼
        self.run_button = ttk.Button(button_frame, text="실행", command=self.run_organizer)
        self.run_button.pack(side=tk.RIGHT, padx=5)
        
        # 종료 버튼
        ttk.Button(button_frame, text="종료", command=self.root.destroy).pack(side=tk.RIGHT, padx=5)
    
    def browse_folder(self):
        folder_path = filedialog.askdirectory(title="과제 폴더 선택")
        if folder_path:
            self.selected_folder.set(folder_path)
    
    def run_organizer(self):
        folder_path = self.selected_folder.get()
        
        if not folder_path:
            messagebox.showerror("오류", "먼저 폴더를 선택해주세요.")
            return
        
        if not os.path.isdir(folder_path):
            messagebox.showerror("오류", "선택한 경로가 유효한 폴더가 아닙니다.")
            return
        
        if self.is_running:
            messagebox.showinfo("알림", "이미 작업이 실행 중입니다. 완료될 때까지 기다려주세요.")
            return
        
        # 실행 버튼 비활성화 및 상태 변경
        self.is_running = True
        self.run_button.config(state=tk.DISABLED)
        
        # 로그 초기화
        self.log_text.delete(1.0, tk.END)
        self.log("작업을 시작합니다...")
        
        # 스레드로 실행
        threading.Thread(target=self.organize_assignment_folder, args=(folder_path,), daemon=True).start()
    
    def log(self, message):
        self.log_queue.put(message)
    
    def update_log(self):
        # 큐에서 로그 메시지를 가져와 텍스트 위젯에 추가
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)  # 스크롤을 최신 내용으로 이동
                self.log_queue.task_done()
        except queue.Empty:
            pass
        
        # 주기적으로 로그 업데이트
        self.root.after(100, self.update_log)
    
    def organize_assignment_folder(self, root_folder):
        """
        과제 폴더를 정리하는 함수
        
        1. "과제" 폴더 아래에 "out" 폴더를 만든다.
        2. "과제" 폴더의 하위 폴더 중에서 폴더명에 "onlinetext"가 포함된 폴더를 모두 삭제한다.
        3. "과제" 폴더의 모든 하위 폴더 안에 있는 파일명을 부모 폴더명의 앞 11글자로 바꾼다.
        4. "과제" 폴더의 하위 폴더 안에 있는 docx 파일을 "out" 폴더로 옮긴다.
        """
        try:
            # 1. "과제" 폴더 아래에 "out" 폴더 생성
            out_folder = os.path.join(root_folder, 'out')
            if not os.path.exists(out_folder):
                os.makedirs(out_folder)
                self.log(f"'out' 폴더가 생성되었습니다: {out_folder}")
            else:
                self.log(f"'out' 폴더가 이미 존재합니다: {out_folder}")
            
            # 2. "onlinetext"가 포함된 하위 폴더 삭제
            for item in os.listdir(root_folder):
                item_path = os.path.join(root_folder, item)
                if os.path.isdir(item_path) and 'onlinetext' in item.lower() and item != 'out':
                    try:
                        shutil.rmtree(item_path)
                        self.log(f"'onlinetext'가 포함된 폴더 삭제: {item_path}")
                    except Exception as e:
                        self.log(f"폴더 삭제 중 오류 발생: {item_path} - {e}")
            
            # 3 & 4. 모든 하위 폴더 내 파일명 변경 및 docx 파일 이동
            for root, dirs, files in os.walk(root_folder):
                # 'out' 폴더는 처리하지 않음
                if 'out' in root.split(os.sep):
                    continue
                
                # 상위 폴더 이름 가져오기 (부모 폴더)
                parent_folder = os.path.basename(root)
                
                # 폴더명이 너무 짧은 경우 처리
                if len(parent_folder) < 11:
                    prefix = parent_folder
                else:
                    prefix = parent_folder[:11]
                
                for file in files:
                    old_file_path = os.path.join(root, file)
                    file_ext = os.path.splitext(file)[1]
                    
                    # 3. 파일명 변경 (부모 폴더명의 앞 11글자 + 확장자)
                    new_filename = f"{prefix}{file_ext}"
                    new_file_path = os.path.join(root, new_filename)
                    
                    try:
                        # 같은 이름의 파일이 이미 존재하는 경우 처리
                        if old_file_path != new_file_path:
                            counter = 1
                            temp_new_path = new_file_path
                            
                            while os.path.exists(temp_new_path):
                                name_without_ext = prefix
                                temp_new_path = os.path.join(root, f"{name_without_ext}_{counter}{file_ext}")
                                counter += 1
                            
                            new_file_path = temp_new_path
                            os.rename(old_file_path, new_file_path)
                            self.log(f"파일명 변경: {old_file_path} -> {new_file_path}")
                        
                        # 4. docx 파일을 out 폴더로 이동
                        if file_ext.lower() == '.docx':
                            # 이동할 파일의 최종 경로
                            dest_path = os.path.join(out_folder, os.path.basename(new_file_path))
                            
                            # 같은 이름의 파일이 이미 존재하는 경우 처리
                            counter = 1
                            temp_dest_path = dest_path
                            
                            while os.path.exists(temp_dest_path):
                                name_without_ext = os.path.splitext(os.path.basename(new_file_path))[0]
                                temp_dest_path = os.path.join(out_folder, f"{name_without_ext}_{counter}{file_ext}")
                                counter += 1
                            
                            dest_path = temp_dest_path
                            shutil.move(new_file_path, dest_path)
                            self.log(f"docx 파일 이동: {new_file_path} -> {dest_path}")
                    
                    except Exception as e:
                        self.log(f"파일 처리 중 오류 발생: {old_file_path} - {e}")
            
            self.log("===== 과제 폴더 정리가 완료되었습니다 =====")
        
        except Exception as e:
            self.log(f"작업 중 오류가 발생했습니다: {e}")
        
        finally:
            # UI 업데이트는 메인 스레드에서 수행해야 함
            self.root.after(0, self.complete_task)
    
    def complete_task(self):
        self.is_running = False
        self.run_button.config(state=tk.NORMAL)
        messagebox.showinfo("완료", "과제 폴더 정리가 완료되었습니다.")

def main():
    root = tk.Tk()
    app = FolderOrganizerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
