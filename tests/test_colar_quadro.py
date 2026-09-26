import unittest
from pathlib import Path

from streamlit.testing.v1 import AppTest

APP = str(Path(__file__).resolve().parent.parent / "app.py")
T = "\t"

# Dados fictícios no formato da página do quadro de inscritos do portal do TRF1
CABECALHO = T.join(["Matrícula", "Nome", "Lotação Atual", "Data de Exercício", "1ª Opção", "2ª Opção"])
LINHAS = [
    T.join(["JU100", "ANA SILVA", "1A VARA DA SJ ALFA/SEÇÃO JUDICIÁRIA DE ALFA", "10/03/2005",
            "9A VARA DA SJ BETA/SEÇÃO JUDICIÁRIA DE BETA - DISPONÍVEL"]),
    T.join(["JU200", "BRUNO LIMA", "2A VARA DA SJ GAMA/SEÇÃO JUDICIÁRIA DE GAMA", "15/06/2008",
            "1A VARA DA SJ ALFA/SEÇÃO JUDICIÁRIA DE ALFA - PRÉ-INSCRIÇÃO"]),
    T.join(["JU300", "CARLA ROCHA", "3A VARA DA SJ DELTA/SEÇÃO JUDICIÁRIA DE DELTA", "20/01/2010",
            "9A VARA DA SJ BETA/SEÇÃO JUDICIÁRIA DE BETA - DISPONÍVEL",
            "5A RELATORIA DA TR/SEÇÃO JUDICIÁRIA DE BETA - PRÉ-INSCRIÇÃO"]),
]
TOPO = [
    "Quadro de Magistrados Inscritos para Remoção - Lista por Antiguidade",
    "Documento: JF 000/2026 - EDITAL de REMOÇÃO de JUIZ FEDERAL - Data de Publicação: 01/01/2026",
    "",
    "Tipo : Somente inscritos",
    "",
]
RODAPE = ["SAU/SUL - Quadra 2, Bloco A - Praça dos Tribunais Superiores", "", "CEP: 70070-900 Brasília/DF", "", "Versão: 1.2.0"]
PAGINA_INTEIRA = (
    ["Tribunal Regional Federal da Primeira Região", "", "Usuário: JU999 - USUÁRIO FICTÍCIO",
     "Lotação: VARA FICTÍCIA", "", "Sistema de Magistrados", "01/01/2026", "Movimentação»", "Dados Pessoais»", "Sair"]
    + TOPO + [CABECALHO] + LINHAS + RODAPE
    + ["", "© 2020 Tribunal Regional Federal da Primeira Região. Todos os direitos reservados."]
)


def colar(texto):
    at = AppTest.from_file(APP, default_timeout=60)
    at.run()
    at.text_area(key="texto_colado").input(texto).run()
    return at


def mensagens(elementos):
    return [e.value for e in elementos]


class TestColarQuadro(unittest.TestCase):
    def test_textos_colados_validos_sao_lidos(self):
        casos = {
            "só a tabela": "\n".join([CABECALHO] + LINHAS),
            "título e rodapé": "\n".join(TOPO + [CABECALHO] + LINHAS + RODAPE),
            "página inteira (Ctrl+A)": "\n".join(PAGINA_INTEIRA),
            "quebras de linha do Windows": "\r\n".join(PAGINA_INTEIRA),
            "menu com tabulações antes da tabela": "\n".join(["Movimentação\tDados Pessoais\tSair"] + TOPO + [CABECALHO] + LINHAS),
        }
        for nome, texto in casos.items():
            with self.subTest(nome):
                at = colar(texto)
                self.assertIn("3 inscritos lidos · 1 vaga disputada.", mensagens(at.info))
                self.assertEqual(len(at.error), 0)

    def test_texto_sem_tabulacoes_mostra_erro(self):
        at = colar("\n".join([CABECALHO] + LINHAS).replace("\t", "    "))
        self.assertEqual(len(at.error), 1)
        self.assertIn("Não encontrei a tabela no texto colado", at.error[0].value)
        self.assertFalse(any("inscritos lidos" in m for m in mensagens(at.info)))

    def test_processamento_a_partir_do_texto_colado(self):
        at = colar("\n".join(PAGINA_INTEIRA))
        next(b for b in at.button if b.label == "Iniciar o processamento").click().run()

        self.assertIn("Análise concluída.", mensagens(at.success))
        quadro = at.dataframe[0].value
        self.assertEqual(quadro["Nome"].tolist(), ["ANA SILVA", "BRUNO LIMA"])
        self.assertEqual(quadro["Destino"].tolist(), [
            "9A VARA DA SJ BETA/SEÇÃO JUDICIÁRIA DE BETA",
            "1A VARA DA SJ ALFA/SEÇÃO JUDICIÁRIA DE ALFA",
        ])


if __name__ == "__main__":
    unittest.main()
