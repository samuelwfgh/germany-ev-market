# Data sources and update policy

## New passenger-car registrations

Primary German reference: Kraftfahrt-Bundesamt (KBA), with VDA/other KBA-derived publication pages used as a current-month cross-check. The long historical series is refreshed from Robbie Andrew's consolidated monthly car-sales dataset because it provides a stable machine-readable history assembled from official national registration data.

The website groups drivetrains as:
- Petrol
- Diesel
- non-plug-in hybrid (HEV)
- plug-in hybrid (PHEV)
- battery electric (BEV)
- Hydrogen
- Other

Registration timing is not identical to retail sales.

## Used-car ownership transfers

Source: KBA FZ 9 (*Besitzumschreibungen von Kraftfahrzeugen und Kraftfahrzeuganhängern*).

The detailed drivetrain chart is updated only when the updater can validate a complete passenger-car drivetrain row. Total used-market volume can therefore be newer than the detailed drivetrain split.

## Petrol and diesel road fuels

Source: BAFA *Amtliche Mineralöldaten*, with historical cross-checks from UFOP/AMI publications.

Units on the site are thousand tonnes per month. These are domestic-delivery/consumption statistics, not literal petrol-station till transactions. Diesel includes its published bio-component treatment and petrol includes bioethanol according to the source table used.

Monthly reports can be delayed and later revised. Missing months remain missing.

## Public charging infrastructure

Source: Bundesnetzagentur Ladesäulenregister / *Ladeinfrastruktur in Zahlen*.

The current headline includes normal and fast public charging points and simultaneous rated charging power when the Bundesnetzagentur publishes it. Historical rows are source snapshots and should not be read as an estimate of private/home charging.

## Hydrogen stations

Current documented operational snapshot: IPHE Germany country update. Historic milestones use documented H2 Mobility/LBST-type releases.

A station can be nominally operational but temporarily unavailable. The dashboard therefore labels this as a documented network snapshot, not real-time availability. There is no invented monthly interpolation between snapshots.

## Fail-closed rule

Every automated source is parsed into a temporary representation first. The updater applies sanity checks before replacing the JSON dataset. If the source is missing, blocked, revised into an unknown layout, or produces implausible values, the previous known-good dataset remains live and `metadata.json` records a warning.
