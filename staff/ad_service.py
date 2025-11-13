# staff/ad_service.py
from django.conf import settings
from ldap3 import Server, Connection, Tls, MODIFY_REPLACE # type: ignore
import ssl
from datetime import datetime, timezone

class ADService:
    def __init__(self):
        config = getattr(settings, 'AD_CONFIG', {})
        self.server = config.get('SERVER')
        self.admin_user = config.get('ADMIN_USER')
        self.admin_password = config.get('ADMIN_PASSWORD')
        self.base_dn = config.get('BASE_DN', 'DC=corp,DC=p-el,DC=ru')
        
    def connect(self):
        """Подключение к AD"""
        try:
            tls = Tls(validate=ssl.CERT_NONE)
            server = Server(self.server, port=636, use_ssl=True, tls=tls)
            return Connection(server, self.admin_user, self.admin_password, auto_bind=True)
        except Exception as e:
            print(f"AD error: {e}")
            return None

    def get_user_status(self, username):
        """Статус учетной записи"""
        conn = self.connect()
        if not conn: return {'error': 'Connection failed'}
        
        try:
            conn.search(self.base_dn, f"(sAMAccountName={username})", 
                       attributes=['userAccountControl', 'lockoutTime', 'displayName'])
            if not conn.entries: return {'found': False}
            
            user = conn.entries[0]
            uac = int(user.userAccountControl.value or 0)
            
            # Упрощенная проверка блокировки
            lockout_time = user.lockoutTime.value
            if lockout_time:
                # Простая проверка: если lockoutTime существует и не равен 0 - заблокирован
                if isinstance(lockout_time, datetime):
                    # Если это datetime и год больше 1601 - заблокирован
                    locked = lockout_time.year > 1601
                else:
                    # Если это число - проверяем что > 0
                    try:
                        locked = int(lockout_time) > 0
                    except:
                        locked = False
            else:
                locked = False
                
            disabled = bool(uac & 2)
            
            return {
                'found': True, 
                'locked': locked, 
                'disabled': disabled,
                'display_name': user.displayName.value if user.displayName else username
            }
        except Exception as e:
            return {'error': str(e)}
        finally:
            conn.unbind()

    def unlock_user(self, username):
        """Разблокировка пользователя"""
        conn = self.connect()
        if not conn: return {'success': False, 'error': 'Connection failed'}
        
        try:
            conn.search(self.base_dn, f"(sAMAccountName={username})", attributes=['distinguishedName'])
            if not conn.entries: return {'success': False, 'error': 'User not found'}
            
            user_dn = conn.entries[0].distinguishedName.value
            conn.modify(user_dn, {'lockoutTime': [(MODIFY_REPLACE, [0])]})
            
            success = conn.result['result'] == 0
            return {'success': success, 'message': f'Пользователь {username} разблокирован'}
        except Exception as e:
            return {'success': False, 'error': str(e)}
        finally:
            conn.unbind()