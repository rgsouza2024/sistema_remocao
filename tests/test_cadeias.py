import unittest

import pandas as pd

from motor_remocao import montar_cadeias, processar_remocao


def inscritos(*linhas):
    df = pd.DataFrame([
        {'Matrícula': mat, 'Nome': nome, 'Lotação Atual': lotacao, 'Lista_Opcoes': opcoes}
        for mat, nome, lotacao, opcoes in linhas
    ])
    return df


class TestMontarCadeias(unittest.TestCase):
    def test_cascata_termina_em_vaga_remanescente(self):
        df = inscritos(
            (1, 'Ana Silva', '1ª Vara Federal de Uberlândia', ['7ª Vara Federal Cível da SJMG - DISPONÍVEL']),
            (2, 'Bia Souza', '2ª Vara Federal de Uberaba', ['1ª Vara Federal de Uberlândia']),
        )
        vagas = ['7ª Vara Federal Cível da SJMG']
        resultado, _, sobras = processar_remocao(df, vagas)

        cadeias = montar_cadeias(resultado, vagas)

        self.assertEqual(len(cadeias), 1)
        self.assertFalse(cadeias[0]['Permuta'])
        self.assertEqual(
            [(e['Unidade'], e['Magistrado']) for e in cadeias[0]['Etapas']],
            [
                ('7ª Vara Federal Cível da SJMG', 'Ana Silva'),
                ('1ª Vara Federal de Uberlândia', 'Bia Souza'),
                ('2ª Vara Federal de Uberaba', None),
            ],
        )
        self.assertEqual(sobras, {'2A VARA FEDERAL DE UBERABA'})

    def test_vaga_do_edital_sem_interessados(self):
        df = inscritos((1, 'Ana Silva', 'Vara A', ['Vara B']))
        resultado, _, _ = processar_remocao(df, ['Vara Z'])

        cadeias = montar_cadeias(resultado, ['Vara Z'])

        self.assertEqual(cadeias, [{'Permuta': False, 'Etapas': [
            {'Unidade': 'Vara Z', 'Magistrado': None, 'Opção Nº': None}
        ]}])

    def test_movimentacoes_fora_do_edital_formam_permuta(self):
        resultado = pd.DataFrame([
            {'Nome': 'Ana', 'Origem': 'Vara Y', 'Vaga': 'VARA X', 'Opção Nº': 1},
            {'Nome': 'Bia', 'Origem': 'Vara X', 'Vaga': 'VARA Y', 'Opção Nº': 1},
        ])

        cadeias = montar_cadeias(resultado, ['Vara V'])

        self.assertEqual(len(cadeias), 2)
        self.assertEqual(cadeias[0]['Etapas'][0]['Magistrado'], None)
        self.assertTrue(cadeias[1]['Permuta'])
        self.assertEqual([e['Magistrado'] for e in cadeias[1]['Etapas']], ['Ana', 'Bia'])


if __name__ == '__main__':
    unittest.main()
