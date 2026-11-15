# 4. Installation and first use

## 4.1 Before you start

1. Verify the release bundle (section 3.4).
2. Back up the workbook you are going to change.
3. Check that your organisation's macro policy allows you to run the code.

## 4.2 Import the modules

Import the six files in this order.

```text
M_STATS_PROBDIST_CORE.bas
M_STATS_PROBDIST_SPECIALFUNCS.bas
M_STATS_PROBDIST_NORMALFAMILY.bas
M_STATS_PROBDIST_TFAMILY.bas
M_STATS_PROBDIST_CONTINUOUS.bas
M_STATS_PROBDIST_DISCRETE.bas
```

1. Open the workbook and press `Alt+F11` to open the VBA editor.
2. Select the intended project in the Project Explorer.
3. Use **File > Import File...** for each `.bas` file.
4. Confirm that the component names match the file names.
5. Choose **Debug > Compile VBAProject**. The project must compile without an error.
6. Save the workbook as `.xlsm`, `.xlsb` or `.xlam`. An `.xlsx` file does not store macros.
7. Close and reopen Excel before the first check.

Use the import command. Do not paste the source into modules you have named yourself: module names and line endings are part of a reproducible installation.

## 4.3 Your first formula

```text
=K_STATS_NormalStandard_Cumulative(1.64485362695147)
```

The result is `0.95`: the probability that a standard normal variable is at most 1.64485362695147. If the cell shows `#NAME?`, see section 11.1.

## 4.4 A five-minute check

Enter these formulas and compare the results.

| Formula | Expected |
|---|---|
| `=K_STATS_NormalStandard_Density(0)` | 0.3989422804 |
| `=K_STATS_NormalStandard_Survival(8)` | 6.220960574E-16 |
| `=K_STATS_NormalStandard_InverseCumulative(0.975)` | 1.959963985 |
| `=K_STATS_Gamma_Mean(3,2)` | 6 |
| `=K_STATS_Binomial_PMF(7,20,0.35)` | 0.1844011864 |
| `=K_STATS_Poisson_InverseCumulative(0.99,25)` | 37 |
| `=K_STATS_Normal_InverseCumulative(1.5,0,1)` | `#NUM!` |
| `=K_STATS_Gamma_Density(0,0.5,1)` | `#NUM!` |

- Format the cells to show enough digits.
- The second result cannot be obtained as `=1-K_STATS_NormalStandard_Cumulative(8)` (section 7.3).
- The last two rows are errors on purpose: 1.5 is not a probability, and the Gamma density with a shape below 1 is unbounded at zero.

If all eight cells match, the installation works.

## 4.5 Optional: run the regression harness

The five-minute check is not a full validation. For a thorough check:

1. Import `M_STATS_PROBDIST_TEST.bas` from the same version.
2. Compile the project.
3. Run `Test_STATS_PROBDIST_RunAll` from the VBA editor.

A successful run reports:

```text
RESULT: ALL TESTS PASSED
```

The harness checks behaviour and error codes. It does not replace the accuracy evidence described in chapter 10.

If your work is subject to review, record the source tag or commit, the files imported, the Excel version and build, the Office bitness, the operating system, and the results of the compile and of these checks.

## 4.6 Upgrading and removing

To upgrade: read the changelog for every version in between, back up the workbook, replace all six modules together, compile, and repeat sections 4.4 and 4.5. Parameterization, error behaviour and accuracy statements can change between versions even when the project still compiles. A locally modified module is a fork: compare it with the old and the new source and reapply your changes deliberately.

To remove: delete the six modules and the test module, compile, and replace the worksheet formulas that still call the deleted functions.
