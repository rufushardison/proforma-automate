"""
circular_solver.py

Resolves circular references between lender fee, interest carry, and total
project cost as used in the Single Tenant Model.

Actual circular dependency (confirmed from template scan):
    base_costs      = all hard + soft line items EXCLUDING lender_fee and interest_carry
    dev_fee         = base_costs * dev_fee_rate          (non-circular, based on base only)
    C32             = base_costs + lender_fee + interest_carry
    total_proj_cost = C32 + dev_fee                      (= C32 * (1 + dev_fee_rate))
    loan_amount     = total_proj_cost * ltc
    lender_fee      = loan_amount * fee_rate             ← circular: depends on total_proj_cost
    interest_carry  = loan_amount * rate * months / 12  ← also circular when months > 0

Algebraic solution (both lender_fee and interest_carry in one pass):
    Let B  = base_costs
    Let r  = dev_fee_rate
    Let L  = ltc
    Let f  = lender_fee_rate
    Let i  = financing_rate
    Let m  = months_of_construction  (for interest carry during construction)
    Let k  = L * (f + i * m/12)      (combined circular coefficient)

    C32    = B + lender_fee + interest_carry
    TC     = C32 * (1 + r)
    loan   = TC * L = C32 * (1+r) * L
    fee    = loan * f
    carry  = loan * i * m/12

    lender_fee + interest_carry = loan * (f + i*m/12) = C32 * (1+r) * L * (f + i*m/12)
    C32 = B + C32 * (1+r) * L * (f + i*m/12)
    C32 * [1 - (1+r)*L*(f + i*m/12)] = B
    C32 = B / [1 - (1+r)*L*(f + i*m/12)]

    Then: TC = C32*(1+r), loan = TC*L, fee = loan*f, carry = loan*i*m/12
"""

from __future__ import annotations


def solve_loan_and_fees(
    total_cost: float,
    fee_rate: float,
) -> dict[str, float]:
    """
    Algebraically resolve loan amount and lender fee.

    Args:
        total_cost:  Total project cost before the lender fee is added (e.g. purchase price
                     + hard costs + soft costs).  This is the base that the lender funds.
        fee_rate:    Lender origination / arrangement fee expressed as a decimal (e.g. 0.01
                     for 1 %).

    Returns:
        dict with keys:
            loan_amount  – gross loan including the rolled-in fee
            lender_fee   – dollar amount of the fee

    Raises:
        ValueError: if fee_rate >= 1 (mathematically undefined / nonsensical)
    """
    if fee_rate >= 1.0:
        raise ValueError(f"fee_rate must be < 1.0, got {fee_rate}")
    if fee_rate < 0.0:
        raise ValueError(f"fee_rate must be >= 0.0, got {fee_rate}")

    loan_amount = total_cost / (1.0 - fee_rate)
    lender_fee = loan_amount * fee_rate
    return {"loan_amount": loan_amount, "lender_fee": lender_fee}


def solve_interest_carry(
    loan_amount: float,
    annual_rate: float,
    hold_months: float,
) -> float:
    """
    Compute interest carry (non-circular once loan_amount is known).

    Args:
        loan_amount:  Gross loan amount (output of solve_loan_and_fees).
        annual_rate:  Annual interest rate as a decimal (e.g. 0.08 for 8 %).
        hold_months:  Projected hold period in months (can be fractional).

    Returns:
        interest_carry as a float (dollar amount).
    """
    return loan_amount * annual_rate * (hold_months / 12.0)


def solve_iterative(
    compute_fn,
    initial_guess: float = 0.0,
    max_iterations: int = 50,
    convergence_threshold: float = 0.01,
) -> float:
    """
    Generic iterative solver for cases where the algebraic form is complex.

    Args:
        compute_fn:            Callable(current_value) -> next_value.
        initial_guess:         Starting value for iteration.
        max_iterations:        Safety cap on iterations.
        convergence_threshold: Stop when |next - current| < this value.

    Returns:
        Converged value.

    Raises:
        RuntimeError: if convergence is not reached within max_iterations.
    """
    value = initial_guess
    for i in range(max_iterations):
        next_value = compute_fn(value)
        if abs(next_value - value) < convergence_threshold:
            return next_value
        value = next_value
    raise RuntimeError(
        f"Iterative solver did not converge after {max_iterations} iterations. "
        f"Last value: {value:.4f}"
    )


def solve_all(
    base_costs: float,
    loan_to_cost: float,
    fee_rate: float,
    annual_rate: float,
    months_of_construction: float,
    dev_fee_rate: float = 0.0,
) -> dict[str, float]:
    """
    High-level entry point — matches the Single Tenant Model circular logic.

    Args:
        base_costs:             All hard + soft costs EXCLUDING lender_fee,
                                interest_carry, and developer_fee.
                                (acquisition + A&E + contingency + closing + etc.)
        loan_to_cost:           LTC as a decimal (e.g. 0.80 for 80%).
        fee_rate:               Lender fee rate as a decimal (e.g. 0.01 for 1%).
        annual_rate:            Annual financing rate as a decimal (e.g. 0.065).
        months_of_construction: Construction / interest-only draw period in months.
                                Used to compute interest carry.  0 for stabilised
                                acquisitions with no carry.
        dev_fee_rate:           Developer fee as a fraction of (base_costs +
                                lender_fee + interest_carry).  Default 0.

    Returns:
        dict with:
            c32             – base_costs + lender_fee + interest_carry (hard+soft total)
            dev_fee         – developer fee dollar amount
            total_proj_cost – c32 + dev_fee
            loan_amount     – total_proj_cost * ltc
            lender_fee      – loan_amount * fee_rate
            interest_carry  – loan_amount * annual_rate * months/12
            equity_required – total_proj_cost - loan_amount
    """
    # Combined circular coefficient: k = (1+r)*L*(f + i*m/12)
    k = (1.0 + dev_fee_rate) * loan_to_cost * (fee_rate + annual_rate * months_of_construction / 12.0)
    if k >= 1.0:
        raise ValueError(
            f"Circular coefficient k={k:.4f} >= 1.0 — check inputs "
            f"(ltc={loan_to_cost}, fee_rate={fee_rate}, rate={annual_rate}, months={months_of_construction})."
        )

    c32 = base_costs / (1.0 - k)
    total_proj_cost = c32 * (1.0 + dev_fee_rate)
    dev_fee = total_proj_cost - c32
    loan_amount = total_proj_cost * loan_to_cost
    lender_fee = loan_amount * fee_rate
    interest_carry = loan_amount * annual_rate * (months_of_construction / 12.0)
    equity_required = total_proj_cost - loan_amount

    return {
        "c32": c32,
        "dev_fee": dev_fee,
        "total_proj_cost": total_proj_cost,
        "loan_amount": loan_amount,
        "lender_fee": lender_fee,
        "interest_carry": interest_carry,
        "equity_required": equity_required,
    }


# ---------------------------------------------------------------------------
# Quick smoke-test (run this file directly to verify)
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    # Reproduce Single Tenant Model sample deal:
    # base_costs = hard+soft excl lender fees and interest carry
    # = 2750000 + 25000 + 75000 + 0 + 25000 + 2500 + 5000 + 125000 + 0 = 3007500
    result = solve_all(
        base_costs=3_007_500,
        loan_to_cost=0.80,
        fee_rate=0.01,
        annual_rate=0.065,
        months_of_construction=0,
        dev_fee_rate=0.04,
    )
    print("Circular solver results (Single Tenant Model sample):")
    for key, v in result.items():
        print(f"  {key:<22} ${v:>15,.2f}")
    print(f"\n  NOTE: Template D22 shows $17,204.13 (stale — macro was never run).")
    print(f"  Financing sheet H8 shows $25,165.54 (uses stale TC, not converged).")
    print(f"  Above are the TRUE converged values after circular resolution.")
