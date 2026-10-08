"""Builds Team04_Manual_Test_Report_Online_Food_Ordering.docx.

Run:  python report-assets/build_report.py   (from the SVV folder)
Figures come from report-assets/slide-N.png (rendered deck slides) and assets/bug-*.png.
"""

from docx_helpers import *  # noqa: F401,F403
from docx_helpers import ROOT

OUT = ROOT / "Team04_Manual_Test_Report_Online_Food_Ordering.docx"

doc = new_document()
cover(doc)
doc.add_page_break()

# ---------------------------------------------------------------- TOC
doc.add_heading("Table of Contents", level=1)
add_field(doc.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u')
doc.add_page_break()

# ---------------------------------------------------------------- 1 intro
doc.add_heading("1. Introduction", level=1)
doc.add_heading("1.1 Purpose", level=2)
para(doc,
     "This report documents the complete manual testing of the Online Food Ordering System (FOODIE). "
     "Test cases were designed with black-box techniques (Boundary Value Analysis, Equivalence Class "
     "Partitioning, Cause–Effect Graphing and Decision Tables) and white-box techniques (control flow "
     "graph, cyclomatic complexity and source code coverage). They were executed across the four "
     "testing levels: unit, integration, system and acceptance. The results, defects and the final "
     "release recommendation are recorded here.")

doc.add_heading("1.2 Application Overview", level=2)
para(doc,
     "FOODIE lets a customer register with a mobile number, search restaurants, add dishes to a cart, "
     "apply a coupon, pay online (UPI / card) and track the order until delivery.")

doc.add_heading("1.3 Scope", level=2)
bullets(doc, [
    ("In scope: ", "Registration & login (auth), menu & search, cart, billing & coupons, payment, order tracking."),
    ("Out of scope: ", "Restaurant/admin portal, real bank settlement, load/stress testing beyond basic "
                       "page-load timing, security penetration testing."),
])

doc.add_heading("1.4 Business Rules Under Test", level=2)
table(doc, ["Rule ID", "Requirement"], [
    ["BR-01", "Mobile number must be exactly 10 digits and start with 6, 7, 8 or 9."],
    ["BR-02", "Quantity of a single dish in the cart must be between 1 and 10 (inclusive)."],
    ["BR-03", "Delivery is free when the item subtotal is ₹199 or more; otherwise a ₹40 fee applies."],
    ["BR-04", "A coupon applies only if the code exists, is not expired and the cart total is ≥ ₹299; "
              "otherwise the matching error message is shown."],
    ["BR-05", "The payable total can never be negative."],
    ["BR-06", "If payment fails, the order must NOT be placed; the user sees “Payment failed” and can retry."],
    ["BR-07", "Pages load in under 3 s; session expires after 15 min idle; layout works on a 360 px mobile screen."],
], widths=[0.9, 5.5], caption="Business rules (requirements) used as the test basis")

doc.add_heading("1.5 Test Environment", level=2)
table(doc, ["Item", "Details"], [
    ["Application build", "FOODIE test build (defect-evidence build)"],
    ["Client", "Google Chrome (mobile view 360×800 and desktop 1920×1080)"],
    ["Operating systems", "Android 14 (mobile), Windows 11 (desktop)"],
    ["Payment", "Sandbox UPI / card gateway (success and decline modes)"],
    ["Unit testing", "Test driver calling billing functions directly; payment gateway stub"],
    ["Evidence", "Device screenshots attached to each defect report"],
], widths=[1.7, 4.7], caption="Test environment")

doc.add_heading("1.6 Test Approach, Entry & Exit Criteria", level=2)
bullets(doc, [
    ("Approach: ", "Requirements were split into 6 scenarios. Test cases were derived with BVA, ECP, "
                   "cause–effect graphing and decision tables (black-box), and with basis-path testing "
                   "(white-box). Each test case is data-driven: one test case runs several test values."),
    ("Entry criteria: ", "Test build deployed, test data and sandbox payment available, requirements frozen."),
    ("Exit criteria: ", "All planned test cases executed, 100% statement and branch coverage of the billing "
                        "function, no open Critical defect, all Major defects fixed or accepted."),
])

doc.add_heading("1.7 Severity & Priority Definitions", level=2)
table(doc, ["Level", "Meaning"], [
    ["Critical", "Breaks a core business flow or causes financial loss; no workaround."],
    ["Major", "Wrong functional behaviour with business impact; a workaround exists."],
    ["Minor", "Cosmetic or message issue with little impact."],
    ["P1 / P2 / P3", "Fix immediately / fix before release / fix when possible."],
], widths=[1.3, 5.1], caption="Severity and priority scale")

# ---------------------------------------------------------------- 2 scenarios
doc.add_heading("2. Test Scenarios", level=1)
para(doc, "Each module the customer interacts with was mapped to one high-level test scenario.")
table(doc, ["Scenario ID", "Module", "Test Scenario", "Linked Test Cases"], [
    ["TS-01", "auth", "Register and log in with a mobile number", "TC-03"],
    ["TS-02", "menu", "Search restaurants and browse the menu", "TC-07, TC-10"],
    ["TS-03", "cart", "Add, update and remove items in the cart", "TC-01, TC-06"],
    ["TS-04", "billing", "Apply coupon and compute the bill total", "TC-02, TC-04, TC-05, TC-06"],
    ["TS-05", "payment", "Check out and pay online", "TC-06, TC-07, TC-08"],
    ["TS-06", "tracking", "Track the order status", "TC-07, TC-10"],
    ["All", "all", "Non-functional: layout, session timeout, page load", "TC-09"],
], widths=[1.0, 0.9, 3.0, 1.5], caption="Test scenarios")
figure(doc, crop_slide(2), "Modules, scenarios and testing pipeline")

# ---------------------------------------------------------------- 3 black box
doc.add_page_break()
doc.add_heading("3. Black-Box Test Case Design", level=1)
para(doc, "Black-box techniques derive test cases from the requirements only, without looking at the code.")

doc.add_heading("3.1 Boundary Value Analysis (BVA)", level=2)
para(doc,
     "Errors cluster at the edges of input ranges. For a range [min, max] we test min−1, min, min+1, "
     "max−1, max and max+1. BVA was applied to BR-02 (quantity 1–10) in TC-01 and to BR-03 "
     "(free-delivery threshold ₹199) in TC-02.")
table(doc, ["Value", "Boundary", "Expected", "Actual", "Result"], [
    ["0", "min − 1", "Rejected", "Rejected", "PASS"],
    ["1", "min", "Accepted", "Accepted", "PASS"],
    ["2", "min + 1", "Accepted", "Accepted", "PASS"],
    ["9", "max − 1", "Accepted", "Accepted", "PASS"],
    ["10", "max", "Accepted", "Accepted", "PASS"],
    ["11", "max + 1", "Rejected", "Accepted (11 in cart)", "FAIL → BUG-01"],
], widths=[0.8, 1.0, 1.4, 1.8, 1.4], status_col=4,
    caption="TC-01: BVA test values for item quantity (valid range 1–10)")
table(doc, ["Subtotal", "Boundary", "Expected fee", "Actual fee", "Result"], [
    ["₹198", "just below", "₹40", "₹40", "PASS"],
    ["₹199", "on boundary", "₹0 (free)", "₹40", "FAIL → BUG-02"],
    ["₹200", "just above", "₹0 (free)", "₹0", "PASS"],
], widths=[0.9, 1.1, 1.4, 1.6, 1.4], status_col=4,
    caption="TC-02: BVA test values for the free-delivery threshold (≥ ₹199)")
figure(doc, crop_slide(3), "Boundary value analysis on quantity and delivery threshold")

doc.add_heading("3.2 Equivalence Class Partitioning (ECP)", level=2)
para(doc,
     "Inputs that the system treats the same way form one equivalence class; one representative value "
     "per class is enough. For BR-01 (mobile number) there is one valid and four invalid classes, "
     "tested together in TC-03.")
table(doc, ["Class", "Type", "Description", "Representative", "Expected", "Actual", "Result"], [
    ["V1", "Valid", "10 digits, starts 6–9", "9876543210", "Accepted", "Accepted", "PASS"],
    ["I1", "Invalid", "Fewer than 10 digits", "98765", "Rejected", "Rejected", "PASS"],
    ["I2", "Invalid", "More than 10 digits", "98765432101", "Rejected", "Rejected", "PASS"],
    ["I3", "Invalid", "Contains letters/symbols", "98765abcde", "Rejected", "Rejected", "PASS"],
    ["I4", "Invalid", "Starts with 0–5", "1234567890", "Rejected", "Rejected", "PASS"],
], widths=[0.5, 0.65, 1.55, 1.15, 0.85, 0.85, 0.65], status_col=6, size=9,
    caption="TC-03: Equivalence classes for the mobile number field")
figure(doc, crop_slide(4), "Equivalence class partitioning of the mobile number")

doc.add_heading("3.3 Cause–Effect Graphing", level=2)
para(doc,
     "The coupon feature (BR-04) depends on several conditions at once, so it was modelled as a "
     "cause–effect graph. Causes are input conditions, effects are system responses.")
table(doc, ["ID", "Cause (input condition)", "ID", "Effect (system response)"], [
    ["C1", "Coupon code exists", "E1", "Discount applied"],
    ["C2", "Coupon is not expired", "E2", "Message “invalid coupon”"],
    ["C3", "Cart total ≥ ₹299", "E3", "Message “coupon expired”"],
    ["", "", "E4", "Message “add ₹X more”"],
], widths=[0.5, 2.6, 0.5, 2.8], caption="Causes and effects for applying a coupon")
para(doc, "Boolean relationships derived from the graph:", bold_lead=None)
code_block(doc, [
    "E1 = C1 AND C2 AND C3",
    "E4 = C1 AND C2 AND NOT C3",
    "E3 = C1 AND NOT C2",
    "E2 = NOT C1",
])
figure(doc, crop_slide(5), "Cause–effect graph for applying a coupon (∧ = AND, dashed ~ = NOT)")

doc.add_heading("3.4 Decision Table", level=2)
para(doc,
     "The graph was converted to a decision table. Three binary conditions give 2³ = 8 combinations; "
     "using don't-care entries (–) they reduce to 4 rules without losing coverage. Each rule is one "
     "test value of TC-04.")
table(doc, ["Conditions / Actions", "R1", "R2", "R3", "R4"], [
    ["C1  Coupon code exists", "F", "T", "T", "T"],
    ["C2  Coupon not expired", "–", "F", "T", "T"],
    ["C3  Cart total ≥ ₹299", "–", "–", "F", "T"],
    ["E1  Apply discount", "", "", "", "✓"],
    ["E2  Show “invalid coupon”", "✓", "", "", ""],
    ["E3  Show “coupon expired”", "", "✓", "", ""],
    ["E4  Show “add ₹X more”", "", "", "✓", ""],
    ["Test data", "FOOD99X, cart ₹350", "SUMMER50 (expired), ₹350", "FEAST100, cart ₹250", "FEAST100, cart ₹350"],
    ["Result", "PASS", "PASS", "PASS", "PASS"],
], widths=[2.0, 1.1, 1.1, 1.1, 1.1], size=9, caption="TC-04: Decision table for applying a coupon")
figure(doc, crop_slide(6), "Decision table derived from the cause–effect graph")

# ---------------------------------------------------------------- 4 white box
doc.add_page_break()
doc.add_heading("4. White-Box Testing & Source Code Coverage", level=1)
para(doc,
     "White-box testing uses the source code. The billing function calculateTotal() was analysed with "
     "a control flow graph (CFG), its cyclomatic complexity was computed, independent paths were "
     "tested (TC-05) and statement, branch and path coverage were measured.")

doc.add_heading("4.1 Source Code Under Test", level=2)
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
])

doc.add_heading("4.2 Control Flow Graph", level=2)
table(doc, ["Node", "Source lines", "Type"], [
    ["A", "L2–L3", "Decision: subtotal <= 199"],
    ["B", "L4", "Add ₹40 delivery fee"],
    ["C", "L6", "Decision: discount > 0"],
    ["D", "L7", "Subtract discount"],
    ["E", "L9", "Decision: total < 0"],
    ["F", "L10", "Reset total to 0"],
    ["G", "L12", "Return total (exit)"],
], widths=[0.8, 1.4, 4.2], caption="CFG nodes")
para(doc, "Edges (9): A→B, A→C, B→C, C→D, C→E, D→E, E→F, E→G, F→G.   Nodes (7): A–G.   "
          "Predicate (decision) nodes (3): A, C, E.")
figure(doc, crop_slide(7), "Source code, control flow graph and cyclomatic complexity")

doc.add_heading("4.3 Cyclomatic Complexity", level=2)
code_block(doc, [
    "Method 1:  V(G) = E − N + 2  = 9 − 7 + 2 = 4",
    "Method 2:  V(G) = P + 1      = 3 + 1     = 4",
    "Method 3:  V(G) = number of regions in the CFG = 4",
])
para(doc, "Therefore at least 4 test cases (independent paths) are needed for full basis-path coverage.")

doc.add_heading("4.4 Independent Paths and Test Data (TC-05)", level=2)
table(doc, ["Path", "Route", "Input (subtotal, discount)", "Expected", "Actual", "Result"], [
    ["P1", "A→C→E→G", "(₹250, ₹0)", "₹250", "₹250", "PASS"],
    ["P2", "A→B→C→E→G", "(₹150, ₹0)", "₹190", "₹190", "PASS"],
    ["P3", "A→C→D→E→G", "(₹300, ₹100)", "₹200", "₹200", "PASS"],
    ["P4", "A→B→C→D→E→F→G", "(₹50, ₹100)*", "₹0", "₹0", "PASS"],
], widths=[0.5, 1.6, 1.5, 0.9, 0.9, 0.8], status_col=5, caption="Basis paths executed in TC-05")
para(doc, "* P4 cannot be reached through the user interface (a coupon requires a cart ≥ ₹299, so the discount "
          "never exceeds the total). It was executed at unit level through a test driver.", italic=True, size=9.5)

doc.add_heading("4.5 Source Code Coverage Problems", level=2)
para(doc, "Executable statements: L2, L3, L4, L6, L7, L9, L10, L12 = 8.   "
          "Branch outcomes: 3 decisions × 2 (true/false) = 6.   Independent paths: V(G) = 4.")
para(doc, "Problem 1: coverage achieved by the black-box (UI) tests alone, i.e. paths P1–P3.", bold_lead=None)
code_block(doc, [
    "Statement coverage = statements executed / total statements × 100",
    "                   = 7 / 8 × 100 = 87.5 %      (L10 never executed)",
    "Branch coverage    = branch outcomes taken / total outcomes × 100",
    "                   = 5 / 6 × 100 = 83.3 %      (E-true never taken)",
    "Path coverage      = paths executed / V(G) × 100",
    "                   = 3 / 4 × 100 = 75 %",
])
para(doc, "Problem 2: coverage after adding the white-box unit test for path P4.")
code_block(doc, [
    "Statement coverage = 8 / 8 × 100 = 100 %",
    "Branch coverage    = 6 / 6 × 100 = 100 %",
    "Path coverage      = 4 / 4 × 100 = 100 %",
])
table(doc, ["Coverage metric", "Black-box only (P1–P3)", "After white-box (P1–P4)"], [
    ["Statement", "7/8 = 87.5%", "8/8 = 100%"],
    ["Branch (decision)", "5/6 = 83.3%", "6/6 = 100%"],
    ["Condition", "5/6 = 83.3%", "6/6 = 100%"],
    ["Basis path", "3/4 = 75%", "4/4 = 100%"],
], widths=[2.0, 2.2, 2.2], caption="Coverage summary for calculateTotal()")
para(doc, "Each decision contains a single atomic condition, so condition coverage equals branch coverage.",
     italic=True, size=9.5)
para(doc,
     "Key observation: 100% path coverage did not reveal BUG-02, because none of the basis-path inputs lies "
     "exactly on ₹199. The defect was found by BVA (TC-02) and its root cause (the <= operator on line 3) was "
     "located by code review during white-box analysis. Black-box and white-box testing complement each other.",
     bold_lead="Note: ")
figure(doc, crop_slide(8), "Basis paths and coverage before and after white-box testing")

# ---------------------------------------------------------------- 5 levels
doc.add_page_break()
doc.add_heading("5. Test Case Design for Testing Levels", level=1)
table(doc, ["Level", "Objective", "Technique / setup", "Test cases", "Result"], [
    ["Unit", "Verify a single function in isolation",
     "White-box basis paths; driver calls the function; payment gateway replaced by a stub", "TC-05", "PASS"],
    ["Integration", "Verify modules work together",
     "Bottom-up: cart + billing → + coupon → + payment; stubs replaced one at a time", "TC-06", "PASS"],
    ["System", "Verify the complete application end to end",
     "Black-box functional, negative and non-functional tests on the full build", "TC-07, TC-08, TC-09",
     "2 PASS / 1 FAIL"],
    ["Acceptance", "Confirm readiness for real customers",
     "Alpha UAT with 4 users outside the team against acceptance criteria", "TC-10", "PASS"],
], widths=[1.0, 1.5, 2.3, 0.9, 0.8], size=9, caption="Test design per testing level")

doc.add_heading("5.1 Unit Testing", level=2)
para(doc,
     "calculateTotal() was tested through all 4 basis paths (TC-05). validateMobile() and applyCoupon() were "
     "also exercised in isolation with the ECP and decision-table data. A test driver supplied inputs and "
     "compared return values; the payment gateway was replaced by a stub that returns a fixed response.")
doc.add_heading("5.2 Integration Testing", level=2)
para(doc,
     "A bottom-up strategy was used: cart and billing were integrated first, then the coupon module, then "
     "payment. TC-06 verified that data passes correctly across interfaces: cart subtotal → bill total → amount "
     "received by the payment module.")
figure(doc, crop_slide(9), "Unit and integration testing")
doc.add_heading("5.3 System Testing", level=2)
para(doc,
     "The full application was tested as a black box: the end-to-end order flow (TC-07), the payment-failure "
     "negative path (TC-08) and non-functional requirements such as layout, session timeout and page load "
     "(TC-09). TC-08 revealed the critical defect BUG-03.")
doc.add_heading("5.4 Acceptance Testing", level=2)
para(doc, "Alpha user acceptance testing (TC-10) used these acceptance criteria:")
bullets(doc, [
    ("AC-1: ", "A first-time user can place an order in under 2 minutes without help."),
    ("AC-2: ", "The bill (items, fee, discount, total) is correct."),
    ("AC-3: ", "Order status updates are shown live until delivery."),
])
para(doc, "All criteria were met. Sign-off is conditional on fixing BUG-03.")
figure(doc, crop_slide(10), "System and acceptance testing")

# ---------------------------------------------------------------- 6 test cases
doc.add_page_break()
doc.add_heading("6. Test Cases", level=1)
doc.add_heading("6.1 Test Case Summary", level=2)
table(doc, ["TC ID", "Title", "Module", "Technique / Level", "Status", "Defect"], [
    ["TC-01", "Item quantity boundaries in cart", "cart", "BVA", "FAIL", "BUG-01"],
    ["TC-02", "Delivery fee at free-delivery threshold", "billing", "BVA", "FAIL", "BUG-02"],
    ["TC-03", "Mobile number validation", "auth", "ECP", "PASS", "–"],
    ["TC-04", "Coupon application rules", "billing", "Cause–effect + decision table", "PASS", "–"],
    ["TC-05", "calculateTotal() basis paths", "billing", "White-box / Unit", "PASS", "–"],
    ["TC-06", "Amount flow cart → bill → payment", "cart, billing, payment", "Integration", "PASS", "–"],
    ["TC-07", "End-to-end order placement", "all", "System", "PASS", "–"],
    ["TC-08", "Order handling when payment fails", "payment", "System (negative)", "FAIL", "BUG-03"],
    ["TC-09", "Layout, session timeout, page load", "all", "System (non-functional)", "PASS", "–"],
    ["TC-10", "User acceptance of ordering flow", "all", "Acceptance (alpha)", "PASS", "–"],
], widths=[0.6, 2.0, 1.0, 1.4, 0.6, 0.7], status_col=4, size=9, caption="Test case summary")

doc.add_heading("6.2 Detailed Test Cases", level=2)

TEST_CASES = [
    {
        "id": "TC-01", "title": "Verify item quantity boundaries in the cart",
        "scenario": "TS-03", "module": "cart", "tech": "Black-box · Boundary Value Analysis (BR-02)",
        "priority": "High", "by": "Bhaskar Lukram",
        "pre": "User is logged in. Restaurant “Burger Corner” is open; “Classic Chicken Burger” (₹219) is available.",
        "steps": ["Open the Burger Corner menu.", "Add “Classic Chicken Burger” to the cart.",
                  "In the cart, type each test value into the quantity field and confirm.",
                  "Observe whether the quantity is accepted and the item total updates."],
        "data": "Quantity = 0, 1, 2, 9, 10, 11",
        "exp": "0 and 11 rejected with “Quantity must be between 1 and 10”; 1, 2, 9, 10 accepted.",
        "act": "0 rejected; 1, 2, 9, 10 accepted; 11 ACCEPTED: 11 items in cart, item total ₹2,409.",
        "status": "FAIL", "defect": "BUG-01",
    },
    {
        "id": "TC-02", "title": "Verify delivery fee at the free-delivery threshold",
        "scenario": "TS-04", "module": "billing", "tech": "Black-box · Boundary Value Analysis (BR-03)",
        "priority": "High", "by": "Bhaskar Lukram",
        "pre": "User is logged in; delivery address “Home, Chennai” saved; cart is empty.",
        "steps": ["Add test-catalogue items so that the item subtotal equals the test value.",
                  "Proceed to checkout.", "Read the delivery fee and “To pay” amount in Bill details."],
        "data": "Subtotal = ₹198, ₹199, ₹200",
        "exp": "₹198 → fee ₹40 (pay ₹238); ₹199 → fee ₹0 (pay ₹199); ₹200 → fee ₹0 (pay ₹200).",
        "act": "₹198 → ₹40; ₹199 → fee ₹40 charged, to pay ₹239; ₹200 → ₹0.",
        "status": "FAIL", "defect": "BUG-02",
    },
    {
        "id": "TC-03", "title": "Verify mobile number validation on registration",
        "scenario": "TS-01", "module": "auth", "tech": "Black-box · Equivalence Class Partitioning (BR-01)",
        "priority": "Medium", "by": "Khushi Raghav",
        "pre": "App opened on the Register screen; user not logged in.",
        "steps": ["Enter the test value in the mobile number field.", "Tap “Send OTP”.",
                  "Observe whether an OTP is sent or a validation error is shown.", "Repeat for each value."],
        "data": "9876543210 (V1), 98765 (I1), 98765432101 (I2), 98765abcde (I3), 1234567890 (I4)",
        "exp": "Only 9876543210 accepted (OTP sent); all others show “Enter a valid 10-digit mobile number”.",
        "act": "As expected for all five values.",
        "status": "PASS", "defect": "–",
    },
    {
        "id": "TC-04", "title": "Verify coupon application rules",
        "scenario": "TS-04", "module": "billing", "tech": "Black-box · Cause–Effect Graph & Decision Table (BR-04)",
        "priority": "High", "by": "Khushi Raghav",
        "pre": "User is logged in. Coupons: FEAST100 (valid, ₹100 off, min ₹299), SUMMER50 (expired 30-Jun-2026).",
        "steps": ["Build a cart with the stated total.", "Open “Apply coupon” and enter the code.",
                  "Tap Apply and observe the message and the bill."],
        "data": "R1: FOOD99X, ₹350 · R2: SUMMER50, ₹350 · R3: FEAST100, ₹250 · R4: FEAST100, ₹350",
        "exp": "R1 “invalid coupon”; R2 “coupon expired”; R3 “add ₹49 more”; R4 ₹100 off, total ₹250.",
        "act": "All four rules behaved as expected.",
        "status": "PASS", "defect": "–",
    },
    {
        "id": "TC-05", "title": "Verify calculateTotal() through all basis paths",
        "scenario": "TS-04", "module": "billing", "tech": "White-box · Basis path testing · Unit level",
        "priority": "High", "by": "Kartik Gupta",
        "pre": "Unit test driver available for calculateTotal(); payment gateway replaced by a stub.",
        "steps": ["Call calculateTotal(subtotal, discount) with each input pair through the driver.",
                  "Compare the returned value with the expected total.", "Record the path covered."],
        "data": "P1 (250, 0) · P2 (150, 0) · P3 (300, 100) · P4 (50, 100)",
        "exp": "₹250 · ₹190 · ₹200 · ₹0",
        "act": "₹250 · ₹190 · ₹200 · ₹0. Coverage: statement, branch and path 100%.",
        "status": "PASS", "defect": "–",
    },
    {
        "id": "TC-06", "title": "Verify amount flows correctly from cart to billing to payment",
        "scenario": "TS-03, TS-04, TS-05", "module": "cart, billing, coupon, payment",
        "tech": "Integration testing (bottom-up)", "priority": "High", "by": "Kartik Gupta",
        "pre": "Cart, billing and coupon modules integrated; payment module integrated with sandbox gateway "
               "that logs the amount received.",
        "steps": ["Add 2 × “Veg Biryani” (₹180 each) to the cart.", "Apply coupon FEAST100.",
                  "Proceed to pay.", "Compare cart subtotal, bill total and amount received by payment."],
        "data": "2 × ₹180, coupon FEAST100",
        "exp": "Subtotal ₹360, discount ₹100, delivery ₹0, total ₹260; payment receives ₹260.",
        "act": "Values matched at every interface.",
        "status": "PASS", "defect": "–",
    },
    {
        "id": "TC-07", "title": "Verify end-to-end order placement and tracking",
        "scenario": "TS-02, TS-05, TS-06", "module": "all", "tech": "System testing (functional)",
        "priority": "High", "by": "Ayush Pandey",
        "pre": "Registered user; sandbox payment set to success.",
        "steps": ["Log in.", "Search “Pizza” and open a restaurant.", "Add an item and apply a valid coupon.",
                  "Pay via UPI.", "Open order tracking and follow the status."],
        "data": "Sandbox UPI: success",
        "exp": "Order confirmed; status moves Placed → Preparing → Out for delivery → Delivered.",
        "act": "As expected.",
        "status": "PASS", "defect": "–",
    },
    {
        "id": "TC-08", "title": "Verify order handling when payment fails",
        "scenario": "TS-05", "module": "payment", "tech": "System testing (negative path, BR-06)",
        "priority": "Critical", "by": "Ayush Pandey",
        "pre": "User logged in; cart worth ₹438; sandbox UPI set to decline.",
        "steps": ["Proceed to pay and choose UPI.", "Let the sandbox decline the transaction.",
                  "Return to the app and open the order status screen."],
        "data": "Amount ₹438, sandbox UPI: decline",
        "exp": "“Payment failed” shown; order not placed / cancelled; “Retry payment” offered.",
        "act": "Order #FO10482 shows “Order Placed” while payment status is “Failed”.",
        "status": "FAIL", "defect": "BUG-03",
    },
    {
        "id": "TC-09", "title": "Verify layout, session timeout and page load",
        "scenario": "All", "module": "all", "tech": "System testing (non-functional, BR-07)",
        "priority": "Medium", "by": "Kartik Gupta",
        "pre": "App available on a 360×800 mobile viewport and on desktop; 4G network throttling enabled.",
        "steps": ["Open home, menu, cart and checkout on the mobile viewport; check for overlap or horizontal scroll.",
                  "Stay idle for 15 minutes, then perform an action.", "Measure home page load time."],
        "data": "Viewport 360×800; idle 15 min; 4G profile",
        "exp": "No layout issues; session expires and asks to log in; load < 3 s.",
        "act": "Layout correct; session expired as required; load time 1.8 s.",
        "status": "PASS", "defect": "–",
    },
    {
        "id": "TC-10", "title": "User acceptance of the ordering flow",
        "scenario": "TS-02 – TS-06", "module": "all", "tech": "Acceptance testing (alpha UAT)",
        "priority": "High", "by": "Ayush Pandey",
        "pre": "4 users from outside the team; acceptance criteria AC-1 to AC-3 agreed.",
        "steps": ["Each user places an order without guidance; time is recorded.",
                  "Bill is checked against a manual calculation.", "Users watch the status updates."],
        "data": "4 users, 1 order each",
        "exp": "AC-1 order < 2 min; AC-2 bill correct; AC-3 live status updates.",
        "act": "Average 1 min 20 s; all bills correct; status updated live.",
        "status": "PASS", "defect": "–",
    },
]

for i, tc in enumerate(TEST_CASES):
    doc.add_heading(f"{tc['id']}: {tc['title']}", level=3)
    kv_table(doc, [
        ("Test Case ID", tc["id"]),
        ("Scenario", tc["scenario"]),
        ("Module", tc["module"]),
        ("Technique / Level", tc["tech"]),
        ("Priority", tc["priority"]),
        ("Preconditions", tc["pre"]),
        ("Test Steps", "\n".join(f"{n}. {s}" for n, s in enumerate(tc["steps"], 1))),
        ("Test Data", tc["data"]),
        ("Expected Result", tc["exp"]),
        ("Actual Result", tc["act"]),
        ("Status", tc["status"]),
        ("Defect ID", tc["defect"]),
        ("Executed by", f"{tc['by']} · Oct 2026"),
    ])

# ---------------------------------------------------------------- 7 results
doc.add_page_break()
doc.add_heading("7. Test Execution Results", level=1)
table(doc, ["Metric", "Value"], [
    ["Test cases planned", "10"],
    ["Test cases executed", "10"],
    ["Passed", "7"],
    ["Failed", "3"],
    ["Pass rate", "70%"],
    ["Defects raised", "3 (1 Critical, 2 Major)"],
], widths=[3.0, 3.4], caption="Execution summary")
table(doc, ["Category", "Test cases", "Executed", "Passed", "Failed"], [
    ["Black-box design (BVA, ECP, CEG/DT)", "TC-01 – TC-04", "4", "2", "2"],
    ["Unit (white-box)", "TC-05", "1", "1", "0"],
    ["Integration", "TC-06", "1", "1", "0"],
    ["System", "TC-07 – TC-09", "3", "2", "1"],
    ["Acceptance", "TC-10", "1", "1", "0"],
    ["Total", "", "10", "7", "3"],
], widths=[2.6, 1.3, 0.8, 0.8, 0.8], caption="Results by technique and testing level")
figure(doc, crop_slide(11), "Execution log and pass/fail results")

# ---------------------------------------------------------------- 8 defects
doc.add_page_break()
doc.add_heading("8. Defect Reports", level=1)
doc.add_heading("8.1 Defect Summary", level=2)
table(doc, ["Defect ID", "Title", "Module", "Linked TC", "Severity", "Priority", "Status"], [
    ["BUG-01", "Quantity 11 accepted when typed manually", "cart", "TC-01", "Major", "P2", "Open"],
    ["BUG-02", "₹40 delivery fee charged at exactly ₹199", "billing", "TC-02", "Major", "P2", "Open"],
    ["BUG-03", "Order shows “Placed” after payment fails", "payment", "TC-08", "Critical", "P1", "Open"],
], widths=[0.75, 2.2, 0.75, 0.75, 0.75, 0.6, 0.6], size=9, caption="Defect summary")

DEFECTS = [
    {
        "id": "BUG-01", "title": "Cart accepts quantity 11 when typed manually",
        "img": BUGS / "bug-01-quantity-11.png", "module": "cart", "tc": "TC-01 (BVA)",
        "sev": "Major", "pri": "P2: fix before release", "by": "Bhaskar Lukram",
        "steps": ["Log in and open “Burger Corner”.", "Add “Classic Chicken Burger” (₹219) to the cart.",
                  "Tap the quantity field and type 11.", "Confirm the value."],
        "exp": "Quantity above 10 is rejected with “Quantity must be between 1 and 10”.",
        "act": "11 items are accepted in the cart; item total ₹2,409 and total ₹2,569. The “+” button stops at "
               "10, but typed input bypasses the limit.",
        "cause": "The upper limit is enforced only by the stepper button; typed input is not validated on the "
                 "client or the server.",
        "fix": "Validate the quantity field on input and on the server; clamp or reject values outside 1–10.",
    },
    {
        "id": "BUG-02", "title": "₹40 delivery fee charged at exactly ₹199",
        "img": BUGS / "bug-02-fee-at-199.png", "module": "billing", "tc": "TC-02 (BVA), confirmed by white-box review",
        "sev": "Major", "pri": "P2: fix before release", "by": "Bhaskar Lukram",
        "steps": ["Log in; set delivery address to Home, Chennai.", "Add items so the item total is exactly ₹199.",
                  "Proceed to checkout and view Bill details."],
        "exp": "Delivery fee FREE (₹0); to pay ₹199.",
        "act": "Delivery fee ₹40; to pay ₹239.",
        "cause": "calculateTotal() line 3 uses “subtotal <= 199” instead of “subtotal < 199” (off-by-one boundary error).",
        "fix": "Change the condition to “subtotal < 199” and add a regression test at ₹198, ₹199 and ₹200.",
    },
    {
        "id": "BUG-03", "title": "Order status shows “Placed” after payment fails",
        "img": BUGS / "bug-03-failed-payment-placed.png", "module": "payment", "tc": "TC-08 (System, negative)",
        "sev": "Critical", "pri": "P1: fix immediately", "by": "Ayush Pandey",
        "steps": ["Log in and build a cart worth ₹438.", "Proceed to pay with UPI.",
                  "Let the transaction fail (sandbox decline).", "Open the order status screen."],
        "exp": "“Payment failed” is shown; the order is cancelled / not placed; “Retry payment” is offered.",
        "act": "Order #FO10482 shows a green “Order Placed” badge while Payment = Failed and Transaction = FAILED; "
               "the screen also says “Payment could not be completed”.",
        "cause": "The order is saved with status “Placed” before the payment result is confirmed; the failure "
                 "callback does not roll the status back.",
        "fix": "Create orders as “Pending payment”; set “Placed” only on a success callback and “Payment failed / "
               "Cancelled” on failure. Add a system test for every payment outcome.",
    },
]

doc.add_heading("8.2 Detailed Defect Reports", level=2)
for d in DEFECTS:
    doc.add_heading(f"{d['id']}: {d['title']}", level=3)
    defect_block(doc, [
        ("Defect ID", d["id"]),
        ("Title", d["title"]),
        ("Module", d["module"]),
        ("Linked Test Case", d["tc"]),
        ("Severity", d["sev"]),
        ("Priority", d["pri"]),
        ("Status", "Open"),
        ("Environment", "FOODIE test build · Chrome on Android 14 (360×800)"),
        ("Reported by / Date", f"{d['by']} · Oct 2026"),
        ("Steps to Reproduce", "\n".join(f"{n}. {s}" for n, s in enumerate(d["steps"], 1))),
        ("Expected Result", d["exp"]),
        ("Actual Result", d["act"]),
        ("Root Cause (analysis)", d["cause"]),
        ("Suggested Fix", d["fix"]),
    ], d["img"], f"{d['id']} screenshot evidence")

doc.add_page_break()

# ---------------------------------------------------------------- 9 RTM
doc.add_heading("9. Requirement Traceability Matrix", level=1)
table(doc, ["Requirement", "Scenario", "Test case(s)", "Result", "Defect"], [
    ["BR-01 Mobile number format", "TS-01", "TC-03", "PASS", "–"],
    ["BR-02 Quantity 1–10", "TS-03", "TC-01", "FAIL", "BUG-01"],
    ["BR-03 Free delivery ≥ ₹199", "TS-04", "TC-02, TC-05", "FAIL", "BUG-02"],
    ["BR-04 Coupon rules", "TS-04", "TC-04, TC-06", "PASS", "–"],
    ["BR-05 Total never negative", "TS-04", "TC-05", "PASS", "–"],
    ["BR-06 No order on failed payment", "TS-05", "TC-08", "FAIL", "BUG-03"],
    ["BR-07 Non-functional", "All", "TC-09", "PASS", "–"],
    ["End-to-end ordering & acceptance", "TS-02, TS-05, TS-06", "TC-07, TC-10", "PASS", "–"],
], widths=[2.2, 1.2, 1.2, 0.8, 0.9], status_col=3, caption="Requirement traceability matrix")

# ---------------------------------------------------------------- 10 conclusion
doc.add_heading("10. Conclusion & Recommendations", level=1)
para(doc,
     "Ten test cases covering six scenarios were designed and executed on the Online Food Ordering System. "
     "Seven passed and three failed (70% pass rate). White-box testing achieved 100% statement, branch and "
     "path coverage of the billing function. Three defects were raised: one Critical and two Major.")
doc.add_heading("10.1 Key Findings", level=2)
bullets(doc, [
    ("Edges hide bugs: ", "Boundary value analysis found 2 of the 3 defects (BUG-01, BUG-02), both exactly at a boundary."),
    ("Code reveals causes: ", "White-box review traced BUG-02 to a single operator and raised coverage from "
                              "87.5% / 83.3% / 75% to 100%."),
    ("Every level counts: ", "The critical BUG-03 appeared only in end-to-end system testing; unit and "
                             "integration tests could not detect it."),
])
doc.add_heading("10.2 Exit Criteria Status", level=2)
table(doc, ["Exit criterion", "Status"], [
    ["All planned test cases executed", "Met (10/10)"],
    ["100% statement and branch coverage of billing function", "Met"],
    ["No open Critical defect", "Not met (BUG-03 open)"],
    ["All Major defects fixed or accepted", "Not met (BUG-01, BUG-02 open)"],
], widths=[4.0, 2.4], caption="Exit criteria")
doc.add_heading("10.3 Recommendation", level=2)
para(doc,
     "The build is NOT READY for release. Fix BUG-03 (P1) and BUG-01 / BUG-02 (P2), then re-execute TC-01, "
     "TC-02 and TC-08 (confirmation testing) and run a full regression of all 10 test cases before sign-off.",
     bold_lead="Verdict: ")

doc.add_heading("10.4 Team Members", level=2)
team_table(doc)

doc.save(OUT)
print(f"Saved {OUT}")
