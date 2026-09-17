# High-Load Outpatient Appointment Scheduling: Discrete-Event Simulation and Policy Optimisation

Discrete-event simulation and policy optimisation for urgent and routine healthcare demand under high utilisation.

## Project overview

This academic case study analyses an appointment-based healthcare operation with urgent and routine patient classes. With demand close to daily capacity, reserving near-term slots for urgent patients can reduce urgent waiting while increasing routine waiting. The project compares implementable scheduling policies and evaluates the trade-off under an overall mean-wait constraint of five days.

## System model

- Independent Poisson arrivals: 3.69 urgent and 2.19 routine patients per working day.
- Six equal appointment positions per working day; each treatment occupies one position.
- Nominal offered load: 5.88 / 6 = 0.98.
- Every arrival receives an initial appointment immediately.
- Waiting time is measured from arrival to treatment start.
- The model excludes no-shows, cancellations, patient preferences, class changes, non-working periods, and overtime.

## Scheduling policies

1. First Available Slot (FAS) benchmark, which is class-blind.
2. Workload-responsive priority policy, which postpones routine bookings according to the current urgent workload.
3. On-call routine (OCR) policy, which protects near-term urgent capacity and releases eligible positions to flexible routine patients.

## Experimental design

- Implemented and validated in AnyLogic Personal Learning Edition.
- 1,000-patient warm-up.
- 50,000 arrivals per run.
- 250 independent replications per policy.
- 95% confidence intervals and relative half-width checks.
- Candidate policies evaluated against an overall mean-wait constraint of five days.

## Key results

| Policy | Parameters | Urgent mean wait (days) | Routine mean wait (days) | Overall mean wait (days) |
| --- | --- | ---: | ---: | ---: |
| FAS | None | 4.148 | 4.148 | 4.148 |
| Workload-responsive | alpha_1 = 0, alpha_2 = 0.35 | 2.123 | 9.053 | 4.703 |
| OCR | eta = 0.2, L = 0, H = 3 | 4.009 | 4.574 | 4.219 |

- The selected workload-responsive policy reduced urgent mean waiting from 4.148 to 2.123 days, a 48.8% reduction relative to FAS.
- Its overall mean wait was 4.703 days, with a 95% confidence interval of [4.513, 4.893], remaining below the five-day constraint.
- The urgent-wait improvement came with a routine mean wait of 9.053 days, highlighting the distributional trade-off.
- The selected OCR policy achieved 97.911% slot utilisation and an overall mean wait of 4.219 days; a rebooked OCR patient saved 0.515 days on average.

## Repository contents

- outputs/zw1f25_MATH6185_submission/ - submitted AnyLogic models and the final report.
- outputs/zw1f25_MATH6185_presentation/ - presentation materials.
- work/ - working materials and supporting documentation.

Open the .alp files in AnyLogic Personal Learning Edition to inspect the submitted simulation models. The report documents the assumptions, policy rules, warm-up procedure, replication design, and confidence interval calculations.

## Limitations

The findings are scenario-based rather than real-world forecasts. They depend on stationary Poisson demand, deterministic service, identical working days, and the absence of cancellations, no-shows, preferences, clinical deterioration, and overtime. The finite parameter search identifies a good tested policy but does not prove global optimality.
