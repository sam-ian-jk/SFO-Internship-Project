import re

def validate_inputs(pno, sno, ss, tp, mv):
    errors = []
    if not re.match(r"^\d{3}-PCA\d{6}-[A-Z]$", pno):
        errors.append("\u274c Invalid Part Number format (e.g., 146-PCA000045-L)")
    if not re.match(r"^AT-\d{4}-\d{2}-\d{4}$", sno):
        errors.append("\u274c Invalid Serial Number format (e.g., AT-5224-13-0003)")
    if not ss:
        errors.append("\u274c Source Supply cannot be empty")
    if not tp:
        errors.append("\u274c Test Point cannot be empty")
    if not re.match(r"^\d+(\.\d+)?[KkMmRr]?$", mv):
        errors.append("\u274c Invalid Measurement format (e.g., 4.49K)")
    return errors

def is_pass(measurements):
    for value in measurements:
        numbers = re.findall(r"[\d.]+", value)
        if not numbers or float(numbers[0]) <= 10:
            return False
    return True
