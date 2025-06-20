import requests
import json
import pandas as pd
import time
TOKEN = "eyJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJQaXBlZnkiLCJpYXQiOjE3NTA0MTk3MTEsImp0aSI6IjkxNTU3NjM4LTAwZmMtNDJlYS05NDRjLWE5NWI3MTk2MGMxZCIsInN1YiI6MzA2NzM5NzE4LCJ1c2VyIjp7ImlkIjozMDY3Mzk3MTgsImVtYWlsIjoiZGVzZW52b2x2aW1lbnRvMkB0NXRlYy5jb20uYnIifX0.nTiOpnKPxH_E0YcMih3leVJUuaW9Fr6cq6Vr_jJ2IZeT61J2-0desMbdKW8O8A2Z_ta6Cgt1BtohpIvtj_x-1Q"  # Obtenha em: Pipefy > Perfil > API Tokens
TABLE_ID = "304107875"
def colocarMascara_cidade(cidade):
    # Se for lista, pega o primeiro elemento
    if isinstance(cidade, list) and cidade:
        cidade = cidade[0]
    # Se for string com colchetes, remove colchetes e aspas
    if isinstance(cidade, str):
        cidade = cidade.strip().replace('[', '').replace(']', '').replace('"', '')
    return cidade.strip()
def colocarMascara_numero(numero):
    if not numero:
        return ""
    numero = str(numero).strip().replace('.', '').replace('-', '')
    if len(numero) == 11:
        return f"({numero[:2]}) {numero[2:7]}-{numero[7:]}"

    elif len(numero) == 10:
        return f"({numero[:2]}) {numero[2:6]}-{numero[6:]}"

    elif len(numero) == 9:
        return f"{numero[:5]}-{numero[5:]}"

    elif len(numero) == 8:
        return f"{numero[:4]}-{numero[4:]}"
   
        
    return numero

def colocarMascara_cnpj(cnpj):
    cnpj = cnpj.strip().replace('.', '').replace('/', '').replace('-', '')
    if len(cnpj) == 14:
        return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
    else:
        return cnpj  # Retorna sem máscara se não tiver 14 dígitos
def consulta_numero_por_cnpj(cnpj):
    
    cnpj_limpo = cnpj.strip().replace('.', '').replace('/', '').replace('-', '')
    url = f"https://minhareceita.org/{cnpj_limpo}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        # O campo de telefone geralmente vem como "telefone" ou "telefone1"
        numero = colocarMascara_numero(data.get("ddd_telefone_1", ""))

        if data.get("ddd_telefone_2"):
            numero += ", " + colocarMascara_numero(data.get("ddd_telefone_2", ""))
        print(f"Consulta realizada para CNPJ: {cnpj} - Telefone: {numero}")
        return numero
    except requests.RequestException as e:
        print(f"Erro ao consultar CNPJ {cnpj}: {e}")
        return None
    
def consulta_socios_por_cnpj(cnpj):
    cnpj_limpo = cnpj.strip().replace('.', '').replace('/', '').replace('-', '')
    url = f"https://minhareceita.org/{cnpj_limpo}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        socios = data.get("qsa", [])
        print(f"Consulta realizada para CNPJ: {cnpj} - Sócios: {socios}")
        return socios
    except requests.RequestException as e:
        print(f"Erro ao consultar CNPJ {cnpj}: {e}")
        return None

def get_all_records():
    url = "https://api.pipefy.com/graphql"
    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Content-Type": "application/json"
    }
    
    all_records = []
    cursor = None
    page_count = 0
    
    while True:
        page_count += 1
        query = {
            "query": f"""
            {{
              table(id: "{TABLE_ID}") {{
                table_records(first: 100, after: {f'"{cursor}"' if cursor else "null"}) {{
                  pageInfo {{
                    hasNextPage
                    endCursor
                  }}
                  edges {{
                    node {{
                      id
                      title
                      created_at
                      record_fields {{
                        name
                        value
                      }}
                    }}
                  }}
                }}
              }}
            }}
            """
        }
        
        response = requests.post(url, json=query, headers=headers)
        data = response.json()
        
        if "errors" in data:
            print("Erro:", json.dumps(data["errors"], indent=2))
            break
            
        if "data" not in data:
            print("Resposta inesperada da API:", json.dumps(data, indent=2, ensure_ascii=False))
            break
        records_data = data["data"]["table"]["table_records"]
        records = [edge["node"] for edge in records_data["edges"]]
        all_records.extend(records)
        
        page_info = records_data["pageInfo"]
        print(f"Página {page_count}: {len(records)} registros")
        
        if not page_info["hasNextPage"]:
            break
            
        cursor = page_info["endCursor"]
        #if( page_count == 2):
        #    break  # Para teste, remova essa linha para obter todos os registros
       

    print(f"\nTotal de registros obtidos: {len(all_records)}")

   

    # Salvar em arquivo JSON
    with open("pipefy_records.json", "w", encoding="utf-8") as f:
        json.dump(all_records, f, ensure_ascii=False, indent=2)

    # Transformar em tabela para Excel apenas com campos desejados
    campos_desejados = [
        'cnpj',
        'cidade'
    ]
    records_flat = []
    consulta_realizada = False
    for record in all_records:
        flat = {}
        flat['Cliente'] = record.get('title', '')
        cnpj_valor = None
        if 'record_fields' in record:
            for field in record['record_fields']:
                nome = field['name'].strip().lower()
                if nome in campos_desejados:
                    if nome == 'cnpj':
                        flat[field['name']] = colocarMascara_cnpj(field['value'])
                        cnpj_valor = field['value']
                    elif nome == 'cidade':
                        flat[field['name']] = colocarMascara_cidade(field['value'])
                    else:
                        flat[field['name']] = field['value']

        # Consulta telefone e sócios pelo CNPJ para cada registro
        if cnpj_valor:
            telefone = consulta_numero_por_cnpj(cnpj_valor)
            socios = consulta_socios_por_cnpj(cnpj_valor)
            flat['Telefones'] = colocarMascara_numero(telefone)
            if socios:
                nomes_socios = [s.get("nome_socio", "") for s in socios]
                flat['sócios'] = ", ".join(nomes_socios)
            else:
                flat['sócios'] = ""
        else:
            flat['Telefones'] = ""
            flat['sócios'] = ""

        records_flat.append(flat)

    df = pd.DataFrame(records_flat)
    df.to_excel("pipefy_records.xlsx", index=False)
    print("Arquivo Excel salvo como pipefy_records.xlsx (apenas campos selecionados)")
    return all_records

# Executar
if __name__ == "__main__":
   get_all_records()