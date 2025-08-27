from langchain.agents import initialize_agent, Tool
from langchain_openai import ChatOpenAI
from tools import consolidar_bases, calcular_vr, validar_resultado
from dotenv import load_dotenv

load_dotenv()
tools = [
    Tool.from_function(
        func=lambda x: consolidar_bases(x),
        name="Consolidar Bases",
        description="Carrega todas as planilhas de um zip ou diretório"
    ),
    Tool.from_function(
        func=lambda x: calcular_vr(),
        name="Calcular VR",
        description="Aplica regra de cálculo do VR para o período enviado"
    ),
    Tool.from_function(
        func=lambda x: validar_resultado(),
        name="Validar Resultado",
        description="Gera relatório de validação dos cálculos"
    )
]

llm = ChatOpenAI(model="gpt-4o-mini", temperature=0)

agent = initialize_agent(
    tools=tools,
    llm=llm,
    agent="zero-shot-react-description",
    verbose=True
)

def main():
     print(agent.run("Carregue as bases do arquivo Desafio 4 - Dados.zip, calcule os benefícios de Maio/2025 e valide o resultado."))

if __name__ == "__main__":
    main()
