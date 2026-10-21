# Data available in the supplied notebooks

Reviewed the saved `LandGorilla.html` and `FCI ApI Intro.html` as code and recorded output, without executing their cells or copying credentials or portfolio rows. These are examples of past responses, not proof of current service availability.

## Land Gorilla

The app uses the same read sequence shown in the notebook:

1. `GET /api/token`, with USER/PASSWORD headers and `api_name=clm` in the request body.
2. `GET /api/clm/loan?fileNumber=rb0<loan number>` to resolve the internal loan ID.
3. `GET /api/clm/pipelineReportTemplates/437/loans/<id>` for the `data.fields` construction attributes.
4. `GET /api/clm/loan/<id>` for balances and the last approved draw date.

| Response | Useful fields |
| --- | --- |
| Pipeline template | `projectCompletedPercentage`, `projectDurationPercentage`, `loanFundedDate`, `lastDrawDate`, `currentLoanDueDate`, `originalLoanDueDate`, `loanStatus`, `loanProgram` |
| Loan summary | `loanBalance`, `balanceToFinish`, `lastApprovedDrawEffectiveDate` |
| Optional context | `propertyAddress`, `city`, `state`, `borrower`, `systemUsers`, `dateTimeLastNote`, `riskLabel` |

`riskLabel` includes ordinary program labels, retainage terms, permit requirements and draw instructions. It is not a severity classification. Schedule elapsed and construction complete also do not establish that work follows a linear schedule. The slide presents these source facts without generating credit alerts from them. `balanceToFinish` is the construction holdback remaining, not necessarily the entire unfunded loan commitment.

The notebook records successful loan-summary and template-437 responses, but also HTTP 500 from a draw-list request and HTTP 403 from budget, transaction, contact and several guessed nested endpoints. A draw-detail request with a known container ID succeeded. The app does not depend on those additional endpoints or probe them during navigation.

Construction data now loads only when **Load construction data** or **Refresh construction data** is clicked for a selected Bridge loan. Navigation never refreshes the portfolio. Saved snapshots retain their retrieval date and time; expired snapshots and failed refreshes are identified inside the collapsed panel.

To check an affected loan locally, using the same client as the app:

```powershell
python tools/diagnose_landgorilla.py --loan 52406
```

The command reads local `.streamlit/secrets.toml` if available, otherwise prompts for credentials without echoing them. It performs read-only requests for one loan and prints endpoint status codes and field-presence booleans. It omits tokens, credentials, loan identifiers, amounts, borrower information and response bodies. That JSON output can be shared to distinguish a lookup failure from a template or balance failure. Do not share the secrets file or notebook authentication cells.

## FCI

The saved examples call GraphQL `getLoanPortfolio` through an authenticated `fci_graphql` helper that is not defined in this HTML export. Queries use `account`, `includeInactive` and `includeUPB`; examples try different identifier filters. Some return three loans while others return none, so an empty response alone does not establish that a loan is absent.

The successful saved tables include `loanAccount`, `lenderAccount`, `prevServiceAccount`, `borrowerFullName`, `currentBalance`, `loanStatus`, `maturityDate`, `nextDueDate`, and property address information.

A broader query requests `investorAssetNumber`, `originatorLoanNumber`, `originalBalance`, `paidToDate`, `daysLate`, `noteRate`, `totalPayment`, boarding/payoff dates and borrower contact fields, but its saved result is empty. Those fields need a successful authenticated response before treating them as available data.

FCI could supply a servicing snapshot: balance, next due date, paid-through date and delinquency. Match its servicing identifiers to a verified workbook deal number before displaying it; do not assume that FCI's internal `loanAccount` equals the workbook Deal Number. A live FCI integration is not added by this change. It needs the existing authenticated helper or a confirmed endpoint/authentication contract and a verified identifier mapping. The current deck continues to label its servicing figures as workbook data.
