"""
create_user.py
──────────────
Tạo password hash để thêm user vào Streamlit Cloud Secrets.

Cách dùng:
    python create_user.py
"""

import hashlib


def main():
    print("=" * 50)
    print("  Tạo tài khoản mới cho Inlinks Filter")
    print("=" * 50)

    username = input("Tên đăng nhập (không dấu, không space): ").strip()
    display  = input("Tên hiển thị (vd: Nguyễn Văn A): ").strip()
    password = input("Mật khẩu: ").strip()

    hashed = hashlib.sha256(password.encode("utf-8")).hexdigest()

    print("\n✅ Thêm đoạn sau vào Streamlit Cloud Secrets:\n")
    print(f'[users.{username}]')
    print(f'name     = "{display}"')
    print(f'password = "{hashed}"')
    print()
    print("─" * 50)
    print("Vào: share.streamlit.io → app → Settings → Secrets")
    print("─" * 50)


if __name__ == "__main__":
    main()
