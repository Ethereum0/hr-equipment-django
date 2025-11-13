# staff/management/commands/test_ad.py
from django.core.management.base import BaseCommand
from django.conf import settings
from ldap3 import Server, Connection, ALL, Tls # type: ignore
import ssl

class Command(BaseCommand):
    help = 'Test AD connection via LDAPS'

    def handle(self, *args, **options):
        self.stdout.write('🔧 Testing AD connection...')
        
        try:
            # Получаем настройки из settings.py
            config = getattr(settings, 'AD_CONFIG', {})
            server = config.get('SERVER', '172.16.15.5')
            port = config.get('PORT', 636)
            admin_user = config.get('ADMIN_USER', '')
            admin_password = config.get('ADMIN_PASSWORD', '')
            base_dn = config.get('BASE_DN', 'DC=corp,DC=p-el,DC=ru')
            
            self.stdout.write(f'Server: {server}:{port}')
            self.stdout.write(f'Admin user: {admin_user}')
            self.stdout.write(f'Base DN: {base_dn}')
            
            if not admin_user or not admin_password:
                self.stdout.write(
                    self.style.ERROR('❌ AD credentials not configured in settings.py')
                )
                return
            
            # Подключение к AD
            tls_config = Tls(validate=ssl.CERT_NONE, version=ssl.PROTOCOL_TLSv1_2)
            server_obj = Server(
                host=server,
                port=port,
                use_ssl=True,
                tls=tls_config,
                get_info=ALL
            )
            
            self.stdout.write('Connecting to AD...')
            conn = Connection(
                server_obj,
                user=admin_user,
                password=admin_password,
                auto_bind=True
            )
            
            # Тестовый поиск
            self.stdout.write('Testing user search...')
            conn.search(
                search_base=base_dn,
                search_filter='(objectClass=user)',
                attributes=['sAMAccountName'],
                size_limit=1
            )
            
            conn.unbind()
            
            self.stdout.write(
                self.style.SUCCESS('✅ AD connection successful!')
            )
            
        except Exception as e:
            self.stdout.write(
                self.style.ERROR(f'❌ Connection failed: {e}')
            )