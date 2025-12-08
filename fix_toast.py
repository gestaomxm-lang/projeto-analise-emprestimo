import re

# Read the backup file
with open('app_test_2_backup.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace st.success with st.toast for specific messages
content = re.sub(
    r'st\.success\("Análise salva no histórico!"\)',
    'st.toast("✅ Análise salva no histórico!", icon="✅")',
    content
)

content = re.sub(
    r'st\.success\("Análise concluída!"\)',
    'st.toast("✅ Análise concluída!", icon="✅")',
    content
)

content = re.sub(
    r'st\.success\("Análise excluída!"\)',
    'st.toast("🗑️ Análise excluída!", icon="🗑️")',
    content
)

content = re.sub(
    r'st\.success\(f"✅ \{len\(analises_selecionadas\)\} análises consolidadas e reanalisadas com sucesso!"\)',
    'st.toast(f"✅ {len(analises_selecionadas)} análises consolidadas e reanalisadas com sucesso!", icon="✅")',
    content
)

content = re.sub(
    r'st\.info\(f"📊 Total de \{len\(df_resultado_consolidado\)\} itens analisados"\)',
    'st.toast(f"📊 Total de {len(df_resultado_consolidado)} itens analisados", icon="📊")',
    content
)

# Write to app_test_2.py
with open('app_test_2.py', 'w', encoding='utf-8') as f:
    f.write(content)

print('File updated successfully!')
