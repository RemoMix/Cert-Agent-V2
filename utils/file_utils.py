import os
import shutil
import hashlib
from datetime import datetime

class FileUtils:
    @staticmethod
    def get_file_hash(file_path):
        """حساب بصمة الملف"""
        hash_md5 = hashlib.md5()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    
    @staticmethod
    def move_to_processed(source_path, dest_dir):
        """نقل الملف إلى مجلد Processed"""
        try:
            if not os.path.exists(dest_dir):
                os.makedirs(dest_dir, exist_ok=True)
            
            filename = os.path.basename(source_path)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # إضافة timestamp لتفادي التعارض
            new_filename = f"{timestamp}_{filename}"
            dest_path = os.path.join(dest_dir, new_filename)
            
            shutil.move(source_path, dest_path)
            return dest_path
        except Exception as e:
            print(f"خطأ في نقل الملف: {str(e)}")
            return None
    
    @staticmethod
    def clean_temp_files(temp_dir, max_age_hours=24):
        """تنظيف الملفات المؤقتة القديمة"""
        try:
            if not os.path.exists(temp_dir):
                return
            
            current_time = datetime.now()
            
            for filename in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, filename)
                
                if os.path.isfile(file_path):
                    file_age = current_time - datetime.fromtimestamp(os.path.getmtime(file_path))
                    
                    if file_age.total_seconds() > max_age_hours * 3600:
                        os.remove(file_path)
                        print(f"تم حذف الملف المؤقت: {filename}")
        
        except Exception as e:
            print(f"خطأ في تنظيف الملفات المؤقتة: {str(e)}")
    
    @staticmethod
    def create_unique_filename(base_name, directory, extension=None):
        """إنشاء اسم فريد للملف"""
        if extension is None:
            name, ext = os.path.splitext(base_name)
        else:
            name = base_name
            ext = extension
        
        counter = 1
        new_name = f"{name}{ext}"
        
        while os.path.exists(os.path.join(directory, new_name)):
            new_name = f"{name}_{counter}{ext}"
            counter += 1
        
        return new_name