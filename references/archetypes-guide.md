# Detailed Slide Archetypes & Layout Blueprints

This reference details the slide-by-slide blueprints for the 4 core medical presentation archetypes supported by `medical-journal-slides`.

## Archetype A: Clinical Trial Plenary Deck (ASCO / ESMO / AHA Style)
- **Slide 1 (Cover - Bookend Dark `#0F172A`)**: Trial title, acronym, lead authors, journal citation, DOI.
- **Slide 2 (Clinical Background - Open 2-Col)**: Disease burden, SoC limitations, biological rationale.
- **Slide 3 (Study Design & Eligibility - Boxed 3-Col `#EFF6FF`)**: Eligibility criteria, randomization arms, endpoints.
- **Slide 4 (Baseline Patient Characteristics - Clean Table)**: Demographics and clinical markers across arms (max 5 cols x 7 rows).
- **Slide 5 (Primary Efficacy Endpoint - Split Hero Stat + Curve)**: `48 pt` Bold HR/CI/p-value callout + cropped Kaplan-Meier curve (300 DPI).
- **Slide 6 (Subgroup & Secondary Analyses - Open + Forest Plot)**: Forest plot + key secondary endpoints (ORR, OS).
- **Slide 7 (Safety & Tolerability - Toxicity Table)**: All-grade vs Grade >= 3 AEs (`#B91C1C` highlight) and discontinuation rates.
- **Slide 8 (Terminal Take-Home Outro - Bookend Dark `#0F172A`)**: Single 6–12 word practice-changing conclusion + DOI rail.

## Archetype B: Hospital Journal Club & Critical Appraisal (IMRaD + CASP)
1. **Slide 1 (Cover - Dark)**: Clinical question + article citation & Impact Factor.
2. **Slide 2 (Clinical Scenario & PICO - Boxed)**: Patient population, Intervention, Comparator, Outcome.
3. **Slide 3 (Methodology & CONSORT Flow - Open + Diagram)**: Blinding, allocation concealment, ITT, sample size power.
4. **Slide 4 (Primary Efficacy Results - Split Hero)**: ARR, RRR, HR/RR, and NNT calculation.
5. **Slide 5 (Secondary Outcomes & Safety - Table)**: Key secondary outcomes and NNH / Grade >= 3 toxicity.
6. **Slide 6 (Internal Validity Appraisal - Open 2-Col)**: Selection, performance, detection, and attrition biases (CASP).
7. **Slide 7 (External Validity & Local Applicability - Open 2-Col)**: Generalizability to local hospital population.
8. **Slide 8 (Terminal Clinical Verdict - Dark)**: Bottom-line verdict on institutional practice adoption.

## Archetype C: Visual Abstract / Quick Take (NEJM / JAMA Style)
- **Single-Frame Visual Abstract (1–3 Slides)**:
  - **Title Band**: Active research finding + journal citation kicker.
  - **Body Band (3 Boxed Pillars)**: Population ($N$, criteria) | Intervention vs Control | Primary Outcome (`44 pt` Bold stat callout).
  - **Footer Band**: One-line clinical conclusion + DOI.

## Archetype D: Systematic Review & Meta-Analysis (PRISMA / Cochrane)
1. **Slide 1 (Cover - Dark)**: Review title, PICO question, PROSPERO registration.
2. **Slide 2 (PRISMA Study Selection - Open + Flowchart)**: Databases searched, records screened, included studies ($k$, total $N$).
3. **Slide 3 (Risk of Bias Assessment - Table/Matrix)**: RoB-2 or Newcastle-Ottawa quality domains.
4. **Slide 4 (Pooled Primary Outcome - Split Hero + Forest Plot)**: Pooled effect estimate, 95% CI, heterogeneity ($I^2$).
5. **Slide 5 (Subgroup, Sensitivity & Publication Bias - Open 2-Col)**: Funnel plot symmetry, Egger test, sensitivity analyses.
6. **Slide 6 (Terminal GRADE Take-Home - Dark)**: GRADE certainty of evidence rating and clinical guideline recommendation.
