from auth_manager import create_user

# Admin
create_user("admin", "Rc2026#@", "Administrador", "admin")

# Gestão
create_user("gestao", "Gestao2026", "Equipe de Gestão", "gestao")

# Unidades (Hospitais)
unidades = [
    ("h_portugal", "HOSPITAL CASA DE PORTUGAL"),
    ("h_menssana", "HOSPITAL CASA MENSSANA"),
    ("h_evangelico", "HOSPITAL CASA EVANGELICO"),
    ("h_laranjeiras", "HOSPITAL CASA RIO LARANJEIRAS"),
    ("h_botafogo", "HOSPITAL CASA RIO BOTAFOGO"),
    ("h_santacruz", "HOSPITAL CASA SANTA CRUZ"),
    ("h_bernardo", "HOSPITAL CASA SAO BERNARDO"),
    ("h_premium", "HOSPITAL CASA PREMIUM"),
    ("h_ilha", "HOSPITAL CASA ILHA DO GOVERNADOR")
]

for user, nome_real in unidades:
    # Senha padrão para unidades
    create_user(user, "Hospital2026", nome_real, "unidade", unit=nome_real)

print("✅ Usuários criados com sucesso!")
