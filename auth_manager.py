import json
import os
import hashlib
import binascii

import remote_persistence

USERS_FILE = "users.json"
_initial_sync_done = False

def ensure_sync():
    global _initial_sync_done
    if not _initial_sync_done:
        print("☁️ [Auth] Buscando banco de dados na nuvem...")
        if remote_persistence.sync_down():
            print("☁️ [Auth] Banco atualizado da nuvem.")
        _initial_sync_done = True

def hash_password(password):
    """Hash password using PBKDF2 with SHA256 and a random salt."""
    salt = hashlib.sha256(os.urandom(60)).hexdigest().encode('ascii')
    pwdhash = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, 100000)
    pwdhash = binascii.hexlify(pwdhash)
    return (salt + pwdhash).decode('ascii')

def verify_password(stored_password, provided_password):
    ensure_sync() # Garante que temos a versão mais recente ao validar senha
    """Verify a stored password against one provided by user"""
    salt = stored_password[:64]
    stored_hash = stored_password[64:]
    pwdhash = hashlib.pbkdf2_hmac('sha256', provided_password.encode('utf-8'), salt.encode('ascii'), 100000)
    pwdhash = binascii.hexlify(pwdhash).decode('ascii')
    return pwdhash == stored_hash

def load_users():
    ensure_sync()
    if not os.path.exists(USERS_FILE):
        return {}
    try:
        with open(USERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

def save_users(users):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, indent=4, ensure_ascii=False)
    # Trigger cloud sync on save
    remote_persistence.sync_up()

def create_user(username, password, name, role, unit=None):
    users = load_users()
    users[username] = {
        "password": hash_password(password),
        "name": name,
        "role": role,
        "unit": unit
    }
    save_users(users)
    return True

def update_password(username, new_password):
    users = load_users()
    if username in users:
        users[username]["password"] = hash_password(new_password)
        save_users(users)
        return True
    return False

def update_user_details(username, name=None, role=None, unit=None):
    users = load_users()
    if username in users:
        if name: users[username]["name"] = name
        if role: users[username]["role"] = role
        if unit is not None: users[username]["unit"] = unit # Pode ser string vazia para limpar
        save_users(users)
        return True
    return False

def delete_user(username):
    users = load_users()
    if username in users:
        del users[username]
        save_users(users)
        return True
    return False
