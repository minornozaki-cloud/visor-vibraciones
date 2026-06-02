# data/ — Datos de entrada

Datos que la aplicación consume.

## Contenido

- **`sap2000/`** — Exportaciones de SAP2000 (Steady-State, Joint Displacements) en `.xlsx` o `.csv`.

## Nota para v2 — barrido de baja frecuencia

Para evaluar correctamente los **transitorios** (partida / parada), el barrido
steady-state de SAP2000 debe **incluir la zona de baja frecuencia (0 → ~50 Hz)**,
no solo la banda alta (40–80 Hz). El cruce resonante del aislador (~3–4 Hz) ocurre
ahí, y sin esos datos la FRF no existe en esa frecuencia.

Formato de columnas esperado:

```
Joint | OutputCase | CaseType | StepType | Freq | U1 | U2 | U3 | R1 | R2 | R3
```

StepTypes: `Mag at Freq` / `Real at Freq` / `Imag at Freq`
