"""Builds Team04_Manual_Test_Report_eDoc.docx, the 11-page version of the report.

Run:  python report-assets/build_report_short.py   (from the SVV folder)
"""

from docx_helpers import *  # noqa: F401,F403
from docx_helpers import ROOT, BUGS

OUT = ROOT / "Team04_Manual_Test_Report_eDoc.docx"
S = 9  # table font size

doc = new_document(body_size=10.5, h_sizes=(14, 12, 11), margins=(0.85, 0.8), space_after=4)
cover(doc, top_gap=4)
doc.add_page_break()

# ---------------------------------------------------------------- 1 intro
doc.add_heading("1. Introduction", level=1)
para(doc,
     "This report documents the manual testing of FOODIE, an Online Food Ordering System where a customer "
     "registers with a mobile number, adds dishes to a cart, applies a coupon, pays online and tracks the "
     "order. Test cases were designed with black-box techniques (BVA, ECP, Cause–Effect Graphing, Decision "
     "Tables) and white-box techniques (control flow graph, cyclomatic complexity, code coverage), and were "
     "executed at the unit, integration, system and acceptance levels.")

doc.add_heading("1.1 Business Rules Under Test", level=2)
table(doc, ["Rule ID", "Requirement"], [
    ["BR-01", "Mobile number must be exactly 10 digits and start with 6, 7, 8 or 9."],
    ["BR-02", "Quantity of a single dish in the cart must be between 1 and 10."],
    ["BR-03", "Delivery is free when the item subtotal is ₹199 or more; otherwise a ₹40 fee applies."],
    ["BR-04", "A coupon applies only if the code exists, is not expired and the cart total is ≥ ₹299."],
    ["BR-05", "The payable total can never be negative."],
    ["BR-06", "If payment fails, the order must NOT be placed; the user can retry."],
    ["BR-07", "Pages load in under 3 s; session expires after 15 min idle; layout works on a 360 px screen."],
], widths=[0.8, 5.75], size=S, caption="Business rules used as the test basis")

doc.add_heading("1.2 Test Scenarios", level=2)
table(doc, ["Scenario ID", "Module", "Test scenario", "Test cases"], [
    ["TS-01", "auth", "Register and log in with a mobile number", "TC-03"],
    ["TS-02", "menu", "Search restaurants and browse the menu", "TC-07, TC-10"],
    ["TS-03", "cart", "Add, update and remove items in the cart", "TC-01, TC-06"],
    ["TS-04", "billing", "Apply coupon and compute the bill total", "TC-02, TC-04, TC-05, TC-06"],
    ["TS-05", "payment", "Check out and pay online", "TC-06, TC-07, TC-08"],
    ["TS-06", "tracking", "Track the order status", "TC-07, TC-10"],
    ["All", "all", "Non-functional: layout, session timeout, page load", "TC-09"],
], widths=[0.95, 0.85, 3.15, 1.6], size=S, caption="Test scenarios")

doc.add_heading("1.3 Environment and Criteria", level=2)
bullets(doc, [
    ("Environment: ", "FOODIE test build on Chrome (Android 14, 360×800 and Windows 11, 1920×1080); sandbox "
                      "UPI/card gateway; unit test driver with a payment stub."),
    ("Entry criteria: ", "Build deployed, test data and sandbox payment available, requirements frozen."),
    ("Exit criteria: ", "All test cases executed, 100% statement and branch coverage of the billing function, "
                        "no open Critical defect, all Major defects fixed or accepted."),
    ("Severity / priority: ", "Critical = core flow broken, no workaround; Major = wrong behaviour, workaround "
                              "exists. P1 = fix immediately; P2 = fix before release."),
])

# ---------------------------------------------------------------- 2 black box
doc.add_page_break()
doc.add_heading("2. Black-Box Test Case Design", level=1)
doc.add_heading("2.1 Boundary Value Analysis (TC-01, TC-02)", level=2)
para(doc, "For a range [min, max] we test min−1, min, min+1, max−1, max and max+1, because errors cluster at "
          "the edges.")
table(doc, ["Quantity", "Boundary", "Expected", "Actual", "Result"], [
    ["0", "min − 1", "Rejected", "Rejected", "PASS"],
    ["1", "min", "Accepted", "Accepted", "PASS"],
    ["2", "min + 1", "Accepted", "Accepted", "PASS"],
    ["9", "max − 1", "Accepted", "Accepted", "PASS"],
    ["10", "max", "Accepted", "Accepted", "PASS"],
    ["11", "max + 1", "Rejected", "Accepted (11 in cart)", "FAIL → BUG-01"],
], widths=[0.9, 1.1, 1.4, 1.8, 1.35], status_col=4, size=S,
    caption="TC-01: item quantity (valid range 1–10)")
table(doc, ["Subtotal", "Boundary", "Expected fee", "Actual fee", "Result"], [
    ["₹198", "just below", "₹40", "₹40", "PASS"],
    ["₹199", "on boundary", "₹0 (free)", "₹40", "FAIL → BUG-02"],
    ["₹200", "just above", "₹0 (free)", "₹0", "PASS"],
], widths=[0.9, 1.1, 1.4, 1.8, 1.35], status_col=4, size=S,
    caption="TC-02: free-delivery threshold (≥ ₹199)")

doc.add_heading("2.2 Equivalence Class Partitioning (TC-03)", level=2)
para(doc, "Inputs treated the same way form one class, so one representative value per class is enough.")
table(doc, ["Class", "Type", "Description", "Value", "Expected", "Actual", "Result"], [
    ["V1", "Valid", "10 digits, starts 6–9", "9876543210", "Accepted", "Accepted", "PASS"],
    ["I1", "Invalid", "Fewer than 10 digits", "98765", "Rejected", "Rejected", "PASS"],
    ["I2", "Invalid", "More than 10 digits", "98765432101", "Rejected", "Rejected", "PASS"],
    ["I3", "Invalid", "Letters or symbols", "98765abcde", "Rejected", "Rejected", "PASS"],
    ["I4", "Invalid", "Starts with 0–5", "1234567890", "Rejected", "Rejected", "PASS"],
], widths=[0.5, 0.65, 1.55, 1.15, 0.85, 0.85, 0.6], status_col=6, size=S,
    caption="TC-03: equivalence classes for the mobile number")

doc.add_heading("2.3 Cause–Effect Graph (TC-04)", level=2)
table(doc, ["ID", "Cause (input condition)", "ID", "Effect (system response)"], [
    ["C1", "Coupon code exists", "E1", "Discount applied"],
    ["C2", "Coupon is not expired", "E2", "Message “invalid coupon”"],
    ["C3", "Cart total ≥ ₹299", "E3", "Message “coupon expired”"],
    ["", "", "E4", "Message “add ₹X more”"],
], widths=[0.5, 2.75, 0.5, 2.8], size=S, caption="Causes and effects for applying a coupon")
code_block(doc, [
    "E1 = C1 AND C2 AND C3        E3 = C1 AND NOT C2",
    "E4 = C1 AND C2 AND NOT C3    E2 = NOT C1",
], size=9)

doc.add_heading("2.4 Decision Table (TC-04)", level=2)
para(doc, "Three conditions give 2³ = 8 combinations; don't-care entries (–) reduce them to 4 rules."
     ).paragraph_format.keep_with_next = True
table(doc, ["Conditions / Actions", "R1", "R2", "R3", "R4"], [
    ["C1  Coupon code exists", "F", "T", "T", "T"],
    ["C2  Coupon not expired", "–", "F", "T", "T"],
    ["C3  Cart total ≥ ₹299", "–", "–", "F", "T"],
    ["E1  Apply discount", "", "", "", "✓"],
    ["E2  “invalid coupon”", "✓", "", "", ""],
    ["E3  “coupon expired”", "", "✓", "", ""],
    ["E4  “add ₹X more”", "", "", "✓", ""],
    ["Test data", "FOOD99X, ₹350", "SUMMER50, ₹350", "FEAST100, ₹250", "FEAST100, ₹350"],
    ["Result", "PASS", "PASS", "PASS", "PASS"],
], widths=[2.15, 1.1, 1.1, 1.1, 1.1], size=S, caption="TC-04: decision table for applying a coupon")
figure(doc, crop_slide(6), "Decision table derived from the cause–effect graph", width=5.2)

# ---------------------------------------------------------------- 3 white box
doc.add_page_break()
doc.add_heading("3. White-Box Testing & Code Coverage", level=1)
para(doc, "The billing function calculateTotal() was analysed with a control flow graph (CFG), its cyclomatic "
          "complexity was computed, all independent paths were tested (TC-05) and coverage was measured.")
code_block(doc, [
    " 1  function calculateTotal(subtotal, discount) {",
    " 2    let total = subtotal;",
    " 3    if (subtotal <= 199) {          // BUG-02: should be  < 199",
    " 4      total = total + 40;           // delivery fee",
    " 5    }",
    " 6    if (discount > 0) {",
    " 7      total = total - discount;",
    " 8    }",
    " 9    if (total < 0) {",
    "10      total = 0;",
    "11    }",
    "12    return total;",
    "13  }",
], size=9)
para(doc, "CFG nodes: A (L2–3, decision), B (L4), C (L6, decision), D (L7), E (L9, decision), F (L10), G (L12, "
          "exit). Edges (9): A→B, A→C, B→C, C→D, C→E, D→E, E→F, E→G, F→G.", size=10)
figure(doc, crop_slide(7), "Source code, control flow graph and cyclomatic complexity", width=5.4)

doc.add_heading("3.1 Cyclomatic Complexity and Basis Paths (TC-05)", level=2)
code_block(doc, [
    "V(G) = E − N + 2 = 9 − 7 + 2 = 4     V(G) = P + 1 = 3 + 1 = 4     Regions = 4",
], size=9)
table(doc, ["Path", "Route", "Input (subtotal, discount)", "Expected", "Actual", "Result"], [
    ["P1", "A→C→E→G", "(₹250, ₹0)", "₹250", "₹250", "PASS"],
    ["P2", "A→B→C→E→G", "(₹150, ₹0)", "₹190", "₹190", "PASS"],
    ["P3", "A→C→D→E→G", "(₹300, ₹100)", "₹200", "₹200", "PASS"],
    ["P4", "A→B→C→D→E→F→G", "(₹50, ₹100)*", "₹0", "₹0", "PASS"],
], widths=[0.5, 1.7, 1.65, 0.9, 0.9, 0.9], status_col=5, size=S, caption="Basis paths executed in TC-05")
para(doc, "* P4 is not reachable from the UI (a coupon needs a cart ≥ ₹299), so it was run through a unit test "
          "driver.", italic=True, size=9)

doc.add_heading("3.2 Code Coverage Problems", level=2)
para(doc, "Statements: 8 (L2, L3, L4, L6, L7, L9, L10, L12). Branch outcomes: 3 decisions × 2 = 6. Paths: V(G) = 4.",
     size=10).paragraph_format.keep_with_next = True
code_block(doc, [
    "Problem 1: black-box tests only (P1–P3)",
    "  Statement = 7/8 × 100 = 87.5 %   (L10 never executed)",
    "  Branch    = 5/6 × 100 = 83.3 %   (E-true never taken)",
    "  Path      = 3/4 × 100 = 75 %",
    "Problem 2: after adding the white-box unit test for P4",
    "  Statement = 8/8 = 100 %   Branch = 6/6 = 100 %   Path = 4/4 = 100 %",
], size=9)
para(doc, "100% path coverage did not reveal BUG-02 because no basis-path input lies exactly on ₹199. BVA found "
          "it (TC-02) and code review located the cause (<= on line 3), so black-box and white-box testing "
          "complement each other.", bold_lead="Note: ", size=10)
figure(doc, crop_slide(8), "Basis paths and coverage before and after white-box testing", width=5.8)

# ---------------------------------------------------------------- 4 levels
doc.add_page_break()
doc.add_heading("4. Test Case Design for Testing Levels", level=1)
table(doc, ["Level", "Objective", "Technique / setup", "Test cases", "Result"], [
    ["Unit", "Verify one function in isolation",
     "White-box basis paths; test driver; payment gateway replaced by a stub", "TC-05", "PASS"],
    ["Integration", "Verify modules work together",
     "Bottom-up: cart + billing → + coupon → + payment; stubs replaced one at a time", "TC-06", "PASS"],
    ["System", "Verify the full application end to end",
     "Functional, negative and non-functional black-box tests on the full build", "TC-07, TC-08, TC-09",
     "2 PASS / 1 FAIL"],
    ["Acceptance", "Confirm readiness for real users",
     "Alpha UAT: 4 users outside the team, against AC-1 to AC-3", "TC-10", "PASS"],
], widths=[0.95, 1.5, 2.45, 0.95, 0.7], size=S, caption="Test design per testing level")
bullets(doc, [
    ("Unit: ", "calculateTotal() through all 4 basis paths; validateMobile() and applyCoupon() with the ECP "
               "and decision-table data."),
    ("Integration: ", "TC-06 checked that cart subtotal → bill total → amount received by payment match."),
    ("System: ", "end-to-end order (TC-07), payment-failure path (TC-08, found BUG-03), layout, timeout and "
                 "page load (TC-09)."),
    ("Acceptance: ", "AC-1 order in under 2 min without help; AC-2 bill correct; AC-3 live status updates. "
                     "All met."),
], size=10)
figure(doc, crop_slide(9), "Unit and integration testing", width=5.0)
figure(doc, crop_slide(10), "System and acceptance testing", width=5.0)

# ---------------------------------------------------------------- 5 test cases
doc.add_page_break()
doc.add_heading("5. Test Cases", level=1)
doc.add_heading("5.1 Test Case Descriptions", level=2)
table(doc, ["TC ID", "Description and steps", "Technique / level", "Module", "Priority", "Executed by"], [
    ["TC-01", "Verify item quantity boundaries: add “Classic Chicken Burger”, type each quantity in the cart, "
              "check it is accepted or rejected.", "BVA (BR-02)", "cart", "High", "Bhaskar Lukram"],
    ["TC-02", "Verify delivery fee at the threshold: build a cart with each subtotal, check the fee in Bill "
              "details.", "BVA (BR-03)", "billing", "High", "Bhaskar Lukram"],
    ["TC-03", "Verify mobile number validation: enter each value on Register, tap “Send OTP”.",
     "ECP (BR-01)", "auth", "Medium", "Khushi Raghav"],
    ["TC-04", "Verify coupon rules: build the cart, apply the code, check message and bill.",
     "Cause–effect graph + decision table (BR-04)", "billing", "High", "Khushi Raghav"],
    ["TC-05", "Verify calculateTotal() through all basis paths using a unit test driver.",
     "White-box / unit", "billing", "High", "Kartik Gupta"],
    ["TC-06", "Verify the amount flows cart → bill → payment: 2 × Veg Biryani, coupon FEAST100, pay.",
     "Integration (bottom-up)", "cart, billing, payment", "High", "Kartik Gupta"],
    ["TC-07", "Verify end-to-end order: log in, search, add item, coupon, pay by UPI, track order.",
     "System (functional)", "all", "High", "Ayush Pandey"],
    ["TC-08", "Verify order handling when payment fails: pay ₹438 by UPI with sandbox set to decline, "
              "open order status.", "System (negative, BR-06)", "payment", "Critical", "Ayush Pandey"],
    ["TC-09", "Verify layout on 360 px, 15-min session timeout and home page load time.",
     "System (non-functional, BR-07)", "all", "Medium", "Kartik Gupta"],
    ["TC-10", "4 users outside the team place an order unaided; time, bill and status checked.",
     "Acceptance (alpha UAT)", "all", "High", "Ayush Pandey"],
], widths=[0.55, 2.6, 1.3, 0.8, 0.6, 0.85], size=8.5, caption="Test case descriptions")

doc.add_heading("5.2 Test Execution and Results", level=2)
table(doc, ["TC ID", "Test data", "Expected result", "Actual result", "Status", "Defect"], [
    ["TC-01", "Qty 0, 1, 2, 9, 10, 11", "0 and 11 rejected; 1–10 accepted", "11 accepted (item total ₹2,409)",
     "FAIL", "BUG-01"],
    ["TC-02", "Subtotal ₹198, ₹199, ₹200", "Fee ₹40, ₹0, ₹0", "Fee ₹40, ₹40, ₹0 (to pay ₹239 at ₹199)",
     "FAIL", "BUG-02"],
    ["TC-03", "V1, I1, I2, I3, I4", "Only V1 accepted", "As expected", "PASS", "–"],
    ["TC-04", "Rules R1–R4", "Invalid / expired / add ₹49 more / ₹100 off", "As expected", "PASS", "–"],
    ["TC-05", "P1–P4", "₹250, ₹190, ₹200, ₹0", "As expected; 100% coverage", "PASS", "–"],
    ["TC-06", "2 × ₹180, FEAST100", "Total ₹260 at every interface", "Values matched", "PASS", "–"],
    ["TC-07", "Sandbox UPI success", "Order confirmed and tracked to delivery", "As expected", "PASS", "–"],
    ["TC-08", "₹438, UPI decline", "“Payment failed”; order not placed", "“Order Placed” with payment Failed",
     "FAIL", "BUG-03"],
    ["TC-09", "360×800, idle 15 min, 4G", "No layout issues; timeout; load < 3 s", "Correct; load 1.8 s",
     "PASS", "–"],
    ["TC-10", "4 users", "Order < 2 min; bill correct; live status", "Avg 1 min 20 s; all correct", "PASS", "–"],
], widths=[0.55, 1.35, 1.65, 1.75, 0.55, 0.85], status_col=4, size=8.5, caption="Execution log")
table(doc, ["Executed", "Passed", "Failed", "Pass rate", "Defects"], [
    ["10", "7", "3", "70%", "3 (1 Critical, 2 Major)"],
], widths=[1.1, 1.1, 1.1, 1.1, 2.15], size=S, caption="Execution summary")

# ---------------------------------------------------------------- 6 defects
doc.add_page_break()
doc.add_heading("6. Defect Reports", level=1)
table(doc, ["Defect ID", "Title", "Module", "Linked TC", "Severity", "Priority", "Status"], [
    ["BUG-01", "Quantity 11 accepted when typed manually", "cart", "TC-01", "Major", "P2", "Open"],
    ["BUG-02", "₹40 delivery fee charged at exactly ₹199", "billing", "TC-02", "Major", "P2", "Open"],
    ["BUG-03", "Order shows “Placed” after payment fails", "payment", "TC-08", "Critical", "P1", "Open"],
], widths=[0.75, 2.2, 0.7, 0.75, 0.75, 0.65, 0.65], size=S, caption="Defect summary")

DEFECTS = [
    ("BUG-01", "Cart accepts quantity 11 when typed manually", "bug-01-quantity-11.png",
     "cart · TC-01 (BVA)", "Major · P2", "Bhaskar Lukram",
     ["Open “Burger Corner” and add “Classic Chicken Burger” (₹219).", "In the cart, type 11 in the quantity "
      "field and confirm."],
     "Rejected with “Quantity must be between 1 and 10”.",
     "11 items accepted; item total ₹2,409, total ₹2,569.",
     "Limit enforced only by the “+” button; typed input not validated.",
     "Validate quantity on input and on the server (1–10)."),
    ("BUG-02", "₹40 delivery fee charged at exactly ₹199", "bug-02-fee-at-199.png",
     "billing · TC-02 (BVA)", "Major · P2", "Bhaskar Lukram",
     ["Add items so the item total is exactly ₹199.", "Go to checkout and view Bill details."],
     "Delivery fee FREE; to pay ₹199.",
     "Delivery fee ₹40; to pay ₹239.",
     "calculateTotal() line 3 uses “<= 199” instead of “< 199”.",
     "Change to “< 199”; add regression tests at ₹198/199/200."),
    ("BUG-03", "Order status shows “Placed” after payment fails", "bug-03-failed-payment-placed.png",
     "payment · TC-08 (System)", "Critical · P1", "Ayush Pandey",
     ["Build a cart worth ₹438 and pay by UPI.", "Let the sandbox decline the payment.",
      "Open the order status screen."],
     "“Payment failed”; order not placed; “Retry payment” offered.",
     "Order #FO10482 shows “Order Placed” while payment is Failed.",
     "Order saved as “Placed” before the payment result; no rollback on failure.",
     "Create orders as “Pending payment”; set “Placed” only on success."),
]

for d_id, title, img, mod, sev, by, steps, exp, act, cause, fix in DEFECTS:
    doc.add_heading(f"{d_id}: {title}", level=3)
    defect_block(doc, [
        ("Defect ID", d_id),
        ("Module / Test case", mod),
        ("Severity / Priority", sev),
        ("Status", "Open"),
        ("Reported by", f"{by} · Oct 2026"),
        ("Steps to reproduce", "\n".join(f"{n}. {s}" for n, s in enumerate(steps, 1))),
        ("Expected result", exp),
        ("Actual result", act),
        ("Root cause", cause),
        ("Suggested fix", fix),
    ], BUGS / img, f"{d_id} screenshot", left_w=4.75, right_w=1.8, img_w=1.45, size=8.5)
figure(doc, crop_slide(12), "Defect log: all three defects with screenshot evidence", width=6.0)

# ---------------------------------------------------------------- 7 conclusion
doc.add_page_break()
doc.add_heading("7. Conclusion", level=1)
para(doc,
     "Ten test cases covering six scenarios were executed: seven passed and three failed (70%). White-box "
     "testing raised coverage of the billing function from 87.5% statement / 83.3% branch / 75% path to 100%. "
     "Three defects were raised: one Critical and two Major.")
bullets(doc, [
    ("Edges hide bugs: ", "BVA found 2 of the 3 defects, both exactly at a boundary."),
    ("Code reveals causes: ", "white-box review traced BUG-02 to a single operator."),
    ("Every level counts: ", "the critical BUG-03 appeared only in end-to-end system testing."),
])
table(doc, ["Requirement", "Test case(s)", "Result", "Defect"], [
    ["BR-01 Mobile number format", "TC-03", "PASS", "–"],
    ["BR-02 Quantity 1–10", "TC-01", "FAIL", "BUG-01"],
    ["BR-03 Free delivery ≥ ₹199", "TC-02, TC-05", "FAIL", "BUG-02"],
    ["BR-04 Coupon rules", "TC-04, TC-06", "PASS", "–"],
    ["BR-05 Total never negative", "TC-05", "PASS", "–"],
    ["BR-06 No order on failed payment", "TC-08", "FAIL", "BUG-03"],
    ["BR-07 Non-functional", "TC-09", "PASS", "–"],
    ["End-to-end ordering & acceptance", "TC-07, TC-10", "PASS", "–"],
], widths=[2.7, 1.6, 1.0, 1.25], status_col=2, size=S, caption="Requirement traceability matrix")
para(doc,
     "NOT READY for release. Fix BUG-03 (P1) and BUG-01 / BUG-02 (P2), re-run TC-01, TC-02 and TC-08, "
     "then run a full regression of all 10 test cases.", bold_lead="Verdict: ")

doc.add_heading("Team Members", level=2)
team_table(doc)

doc.save(OUT)
print(f"Saved {OUT}")
