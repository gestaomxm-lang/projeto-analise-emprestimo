import base64

caminho_exe = "baixar_teste2.exe"
caminho_saida = "exe_embutido.py"

with open(caminho_exe, "rb") as f:
    dados = f.read()

# Converte para Base64 (texto)
b64 = base64.b64encode(dados).decode("utf-8")

with open(caminho_saida, "w", encoding="utf-8") as f:
    f.write("import base64\n\n")
    f.write("exe_b64 = '''\\\n")
    f.write(b64)
    f.write("'''\n\n")
    f.write(
        "with open('baixar_teste2_restaurado.exe', 'wb') as f_out:\n"
        "    f_out.write(base64.b64decode(exe_b64))\n"
        "print('Executável restaurado como baixar_teste2_restaurado.exe')\n"
    )

print("Arquivo exe_embutido.py criado com sucesso.")