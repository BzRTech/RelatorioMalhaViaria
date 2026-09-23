# -*- coding: utf-8 -*-
"""
Executa um relatório num PROCESSO SEPARADO, que termina ao final.

pandas/matplotlib não devolvem a memória ao sistema entre execuções: no mesmo
processo, cada relatório somaria centenas de MB até o servidor (512 MB no
plano free do Render) matar o worker com erro 502. Um processo novo por
relatório libera tudo ao terminar.

Uso interno:  python tarefa.py <entrada> <municipio> <saida> [logo]
Escreve o resultado em <saida>/resultado.json.
"""

import json
import os
import subprocess
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


class ProcessoInterrompido(RuntimeError):
    """O processo filho morreu sem responder (ex.: falta de memória)."""


def gerar_isolado(entrada, municipio, saida, logo_path=None, timeout=280):
    """Roda o relatório num processo Python novo e devolve o resumo."""
    os.makedirs(saida, exist_ok=True)
    arq_resultado = os.path.join(saida, "resultado.json")
    cmd = [sys.executable, os.path.join(BASE_DIR, "tarefa.py"),
           entrada, municipio, saida, logo_path or ""]
    env = dict(os.environ, PYTHONIOENCODING="utf-8", MPLBACKEND="Agg")
    proc = subprocess.run(cmd, cwd=BASE_DIR, env=env, timeout=timeout,
                          capture_output=True, text=True, encoding="utf-8", errors="replace")
    if not os.path.exists(arq_resultado):
        raise ProcessoInterrompido(
            f"código {proc.returncode}: {(proc.stderr or '').strip()[-800:]}")
    with open(arq_resultado, encoding="utf-8") as fh:
        res = json.load(fh)
    os.remove(arq_resultado)
    if "erro" in res:
        raise RuntimeError(res["erro"])
    return res


def _executar(entrada, municipio, saida, logo_path):
    import relatorio_malha_viaria as R

    res = R.executar(entrada, municipio, saida, logo_path=logo_path or None,
                     gerar_pptx=True, log=lambda *a: None)
    rel = res["rel"]

    resumo = [("Trechos", R.formatar_inteiro_br(rel.get("total_trechos", rel["total_registros"])))]
    if rel.get("total_vias"):
        resumo.append(("Logradouros", R.formatar_inteiro_br(rel["total_vias"])))
    if "total_km" in rel:
        resumo.append(("Extensão", f"{R.formatar_numero_br(rel['total_km'], 2)} km"))
    if "pct_sem_nome_vias" in rel:
        resumo.append(("Sem denominação", R.formatar_pct(rel["pct_sem_nome_vias"], 1)))
    return {"data": rel["data"], "resumo": resumo, "arquivos": res["arquivos"]}


if __name__ == "__main__":
    entrada, municipio, saida, logo = sys.argv[1:5]
    try:
        resultado = _executar(entrada, municipio, saida, logo)
    except Exception as e:  # noqa: BLE001 - erro vai para o site mostrar
        resultado = {"erro": str(e)}
    with open(os.path.join(saida, "resultado.json"), "w", encoding="utf-8") as fh:
        json.dump(resultado, fh, ensure_ascii=False)
