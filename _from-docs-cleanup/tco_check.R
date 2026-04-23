# Iteration: 2
# TCO v3.2 parameter check — run standalone

analysis_period   <- 5
annual_km         <- 12563
elec_escalation   <- 0.02
fuel_escalation   <- 0.00
ins_escalation    <- 0.025
maint_escalation  <- 0.025

ev_price     <- 39950
ev_residual  <- round(ev_price * 0.3)
icev_price   <- 28000
icev_residual <- round(icev_price * 0.3)
ev_own_annual   <- (ev_price - ev_residual) / analysis_period
icev_own_annual <- (icev_price - icev_residual) / analysis_period

ev_consumption  <- 14.39
charging_eff    <- 0.9
elec_rate_commercial <- 0.544
ev_energy_yr1   <- (annual_km / 100) * ev_consumption * elec_rate_commercial

icev_consumption <- 8.0
fuel_price       <- 2.21
icev_fuel_yr1    <- (annual_km / 100) * icev_consumption * fuel_price

ev_maint_yr1   <- 1000
icev_maint_yr1 <- 1339.705

ins_rate       <- 2800 / icev_price
ev_ins_yr1     <- round(ev_price * ins_rate)
icev_ins_yr1   <- 2800

ev_infra_annual <- 360

ev_lease_annual   <- 975 * 12
icev_lease_annual <- 833 * 12

years <- 1:analysis_period

cat("=== KEY PARAMETERS ===\n")
cat(sprintf("EV energy yr1 (no double count): Afl %.0f\n", ev_energy_yr1))
cat(sprintf("ICEV fuel yr1 (8.0 L/100km): Afl %.0f\n", icev_fuel_yr1))
cat(sprintf("Energy savings yr1: Afl %.0f\n", icev_fuel_yr1 - ev_energy_yr1))
cat(sprintf("EV insurance yr1: Afl %.0f\n", ev_ins_yr1))
cat(sprintf("ICEV insurance yr1: Afl %.0f\n", icev_ins_yr1))
cat(sprintf("EV lease annual: Afl %.0f\n", ev_lease_annual))
cat(sprintf("ICEV lease annual: Afl %.0f\n", icev_lease_annual))
cat(sprintf("Lease premium: Afl %.0f/yr\n", ev_lease_annual - icev_lease_annual))

cat("\n=== PURCHASE MODE ===\n")
ev_purchase_total <- 0
icev_purchase_total <- 0
for (y in years) {
  ev_yr <- ev_own_annual + ev_energy_yr1*(1+elec_escalation)^(y-1) + ev_maint_yr1*(1+maint_escalation)^(y-1) + ev_ins_yr1*(1+ins_escalation)^(y-1) + ev_infra_annual
  icev_yr <- icev_own_annual + icev_fuel_yr1*(1+fuel_escalation)^(y-1) + icev_maint_yr1*(1+maint_escalation)^(y-1) + icev_ins_yr1*(1+ins_escalation)^(y-1)
  ev_purchase_total <- ev_purchase_total + ev_yr
  icev_purchase_total <- icev_purchase_total + icev_yr
}
cat(sprintf("EV 5-yr TCO: Afl %s\n", format(round(ev_purchase_total), big.mark=",")))
cat(sprintf("ICEV 5-yr TCO: Afl %s\n", format(round(icev_purchase_total), big.mark=",")))
pct_p <- (ev_purchase_total - icev_purchase_total) / icev_purchase_total * 100
cat(sprintf("EV is %.1f%% %s than ICEV\n", abs(pct_p), ifelse(pct_p > 0, "MORE expensive", "CHEAPER")))

cat("\n=== LEASE MODE ===\n")
ev_lease_total <- 0
icev_lease_total <- 0
for (y in years) {
  ev_yr <- ev_lease_annual + ev_energy_yr1*(1+elec_escalation)^(y-1) + ev_ins_yr1*(1+ins_escalation)^(y-1) + ev_infra_annual
  icev_yr <- icev_lease_annual + icev_fuel_yr1*(1+fuel_escalation)^(y-1) + icev_ins_yr1*(1+ins_escalation)^(y-1)
  ev_lease_total <- ev_lease_total + ev_yr
  icev_lease_total <- icev_lease_total + icev_yr
}
cat(sprintf("EV 5-yr TCO: Afl %s\n", format(round(ev_lease_total), big.mark=",")))
cat(sprintf("ICEV 5-yr TCO: Afl %s\n", format(round(icev_lease_total), big.mark=",")))
pct_l <- (ev_lease_total - icev_lease_total) / icev_lease_total * 100
cat(sprintf("EV is %.1f%% %s than ICEV\n", abs(pct_l), ifelse(pct_l > 0, "MORE expensive", "CHEAPER")))

cat("\n=== BREAK-EVEN KM (LEASE) ===\n")
for (km in seq(5000, 50000, by=500)) {
  ev_e <- (km / 100) * ev_consumption * elec_rate_commercial
  icev_e <- (km / 100) * icev_consumption * fuel_price
  savings <- icev_e - ev_e
  net <- savings - (ev_lease_annual - icev_lease_annual) - (ev_ins_yr1 - icev_ins_yr1) - ev_infra_annual
  if (net >= 0) {
    cat(sprintf("Break-even at approximately %s km/year\n", format(km, big.mark=",")))
    break
  }
}

cat("\n=== COST PER KM ===\n")
cat(sprintf("Purchase: EV Afl %.2f/km, ICEV Afl %.2f/km\n",
    ev_purchase_total / (annual_km * analysis_period),
    icev_purchase_total / (annual_km * analysis_period)))
cat(sprintf("Lease: EV Afl %.2f/km, ICEV Afl %.2f/km\n",
    ev_lease_total / (annual_km * analysis_period),
    icev_lease_total / (annual_km * analysis_period)))
