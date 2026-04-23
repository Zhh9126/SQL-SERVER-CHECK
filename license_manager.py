# license_manager.py
import os
import json
import hashlib
import base64
import time
from datetime import datetime, timedelta

# Python 3.6兼容的加密实现
class SimpleCrypto:
    """简单的加密解密类，兼容Python 3.6以上"""
    
    def __init__(self, secret="ODB_SECRET_2024"):
        self.secret = secret.encode('utf-8')
    
    def _xor_encrypt_decrypt(self, data, key):
        """使用XOR进行加密/解密"""
        key_len = len(key)
        result = bytearray()
        for i, byte in enumerate(data):
            result.append(byte ^ key[i % key_len])
        return bytes(result)
    
    def encrypt(self, text):
        """加密文本"""
        data = text.encode('utf-8')
        # 生成密钥
        key = hashlib.sha256(self.secret).digest()[:32]
        # XOR加密
        encrypted = self._xor_encrypt_decrypt(data, key)
        # Base64编码
        return base64.b64encode(encrypted).decode('utf-8')
    
    def decrypt(self, token):
        """解密文本"""
        try:
            # Base64解码
            encrypted = base64.b64decode(token.encode('utf-8'))
            # 生成密钥
            key = hashlib.sha256(self.secret).digest()[:32]
            # XOR解密
            decrypted = self._xor_encrypt_decrypt(encrypted, key)
            return decrypted.decode('utf-8')
        except Exception as e:
            raise ValueError(f"解密失败: {e}")

class LicenseValidator:
    """严格的许可证验证类 - 防止删除文件重置试用期"""

    def __init__(self):
        self.license_file = "mss_inspector.lic"
        self.trial_days = 36500
        self.crypto = SimpleCrypto()
        self._init_license_system()

    def _parse_datetime(self, date_string):
        """解析日期字符串，兼容Python 3.6"""
        try:
            # 尝试Python 3.7+的fromisoformat
            return datetime.fromisoformat(date_string)
        except AttributeError:
            # Python 3.6兼容：手动解析ISO格式
            try:
                # 格式: 2024-01-01T10:30:00
                if 'T' in date_string:
                    date_part, time_part = date_string.split('T')
                    year, month, day = map(int, date_part.split('-'))
                    time_parts = time_part.split(':')
                    hour, minute = int(time_parts[0]), int(time_parts[1])
                    second = int(time_parts[2].split('.')[0]) if len(time_parts) > 2 else 0
                    return datetime(year, month, day, hour, minute, second)
                else:
                    # 格式: 2024-01-01 10:30:00
                    date_part, time_part = date_string.split(' ')
                    year, month, day = map(int, date_part.split('-'))
                    hour, minute, second = map(int, time_part.split(':'))
                    return datetime(year, month, day, hour, minute, second)
            except Exception as e:
                print(f"日期解析错误: {e}")
                return datetime.now()

    def _format_datetime(self, dt):
        """格式化日期为字符串，兼容Python 3.6"""
        return dt.strftime('%Y-%m-%dT%H:%M:%S')

    def _init_license_system(self):
        """初始化许可证系统"""
        # 创建必要的目录和文件
        if not os.path.exists(self.license_file):
            self._create_trial_license()

    def _create_trial_license(self):
        """创建试用许可证"""
        create_time = datetime.now()
        expire_time = create_time + timedelta(days=self.trial_days)
        
        license_data = {
            "type": "TRIAL",
            "create_time": self._format_datetime(create_time),
            "expire_time": self._format_datetime(expire_time),
            "machine_id": self._get_machine_id(),
            "signature": self._generate_signature("TRIAL")
        }
        encrypted_data = self.crypto.encrypt(json.dumps(license_data))
        with open(self.license_file, 'w') as f:
            f.write(encrypted_data)

    def _get_machine_id(self):
        """获取机器标识（简化版）"""
        try:
            import platform
            import socket
            machine_info = f"{platform.node()}-{platform.system()}-{platform.release()}"
            return hashlib.md5(machine_info.encode()).hexdigest()[:16]
        except:
            return "unknown_machine"

    def _generate_signature(self, license_type):
        """生成许可证签名"""
        key = "ODB2024TRL"  # 试用版密钥
        signature_data = f"{license_type}-{datetime.now().strftime('%Y%m%d')}-{key}"
        return hashlib.sha256(signature_data.encode()).hexdigest()

    def _verify_signature(self, license_data):
        """验证许可证签名"""
        try:
            expected_signature = self._generate_signature(license_data["type"])
            return license_data["signature"] == expected_signature
        except:
            return False

    def validate_license(self):
        """验证许可证有效性"""
        try:
            # 检查许可证文件
            if not os.path.exists(self.license_file):
                return False, "许可证文件不存在", 0

            # 读取并解密许可证
            with open(self.license_file, 'r') as f:
                encrypted_data = f.read().strip()
            
            decrypted_data = self.crypto.decrypt(encrypted_data)
            license_data = json.loads(decrypted_data)

            # 验证签名
            if not self._verify_signature(license_data):
                return False, "许可证签名无效", 0

            # 检查过期时间
            expire_time = self._parse_datetime(license_data["expire_time"])
            remaining_days = (expire_time - datetime.now()).days

            if remaining_days < 0:
                return False, "许可证已过期", 0

            license_type = license_data.get("type", "TRIAL")
            
            if license_type == "TRIAL":
                return True, f"试用版许可证有效，剩余 {remaining_days} 天", remaining_days
            else:
                return True, f"{license_type}版许可证有效", remaining_days

        except Exception as e:
            return False, f"许可证验证失败: {str(e)}", 0

    def get_license_info(self):
        """获取许可证信息"""
        try:
            if not os.path.exists(self.license_file):
                return {"type": "TRIAL", "status": "未找到许可证文件"}
            
            with open(self.license_file, 'r') as f:
                encrypted_data = f.read().strip()
            
            decrypted_data = self.crypto.decrypt(encrypted_data)
            license_data = json.loads(decrypted_data)
            
            expire_time = self._parse_datetime(license_data["expire_time"])
            remaining_days = (expire_time - datetime.now()).days
            
            license_data["remaining_days"] = remaining_days
            license_data["status"] = "有效" if remaining_days > 0 else "已过期"
            
            return license_data
            
        except Exception as e:
            return {"type": "UNKNOWN", "status": f"错误: {str(e)}"}
