# Imports the Python file.
import voting as voting

# Imports the openpyxl library.
import openpyxl

# Opens the voting_2.xlsx workbook.
workbook = openpyxl.load_workbook("voting_2.xlsx")

# Selects the sheet from the workbook.
values = workbook.active

def test_1_generate_preferences():
    expected_output = {1: [4, 2, 1, 3], 2: [4, 3, 1, 2], 3: [4, 3, 1, 2], 4: [1, 3, 4, 2], 5: [2, 3, 4, 1], 6: [2, 1, 3, 4]}
    print("Expected Output:", expected_output)
    actual_output = voting.generate_preferences(values)
    print("Actual Output:", actual_output)
    print("Correct")

# test_1_generate_preferences()

# STV Testing

def test_1_STV():
    expected_output = 3
    print("Expected Output:", expected_output)
    actual_output = voting.STV(stv_test_1_values, 3)
    print("Actual Output:", actual_output)
    
stv_test_1_values = {1: [3, 2, 1, 4],
                     2: [4, 1, 2, 3],
                     3: [3, 2, 1, 4],
                     4: [4, 1, 2, 3]}

# test_1_STV()

stv_test_2_values = {1: [4, 3, 1, 2],
                     2: [3, 4, 2, 1],
                     3: [4, 1, 3, 2],
                     4: [3, 2, 1, 4]}

def test_2a_STV():
    expected_output = 4
    print("Expected Output:", expected_output)
    actual_output = voting.STV(stv_test_2_values, "max")
    print("Actual Output:", actual_output)
    
test_2a_STV()

def test_2b_STV():
    expected_output = 3
    print("Expected Output:", expected_output)
    actual_output = voting.STV(stv_test_2_values, "min")
    print("Actual Output:", actual_output)

# test_2b_STV()

def test_2c_STV():
    expected_output = 4
    print("Expected Output:", expected_output)
    actual_output = voting.STV(stv_test_2_values, 1)
    print("Actual Output:", actual_output)

# test_2c_STV()

def test_2d_STV():
    expected_output = 3
    print("Expected Output:", expected_output)
    actual_output = voting.STV(stv_test_2_values, 2)
    print("Actual Output:", actual_output)

# test_2d_STV()
    
# Scoring Rule Testing
    
scoring_rule_test_1_values = {1: [4, 1, 2, 3],
                              2: [3, 4, 1, 2],
                              3: [2, 3, 4, 1],
                              4: [1, 2, 3, 4]}

score_vector = [0.5, 0.5, 0.5, 0.5]

def test_1_scoring_rule():
    expected_output = 3
    print("Expected Output:", expected_output)
    actual_output = voting.scoring_rule(scoring_rule_test_1_values, score_vector, 2)
    print("Actual Output:", actual_output)

# test_1_scoring_rule()
    
scoring_rule_test_2_values = {1: [1, 2, 3, 4],
                              2: [2, 3, 4, 1],
                              3: [3, 4, 1, 2],
                              4: [4, 1, 2, 3]}

def test_2_scoring_rule():
    expected_output = 1
    print("Expected Output:", expected_output)
    actual_output = voting.scoring_rule(scoring_rule_test_2_values, score_vector, 1)
    print("Actual Output:", actual_output)

# test_2_scoring_rule()

score_vector_2 = [0, 0, 0, 0]
    
def test_3_scoring_rule():
    expected_output = 4
    print("Expected Output:", expected_output)
    actual_output = voting.scoring_rule(scoring_rule_test_2_values, score_vector_2, "max")
    print("Actual Output:", actual_output)

# test_3_scoring_rule()

plurality_test_1_values = {1: [1, 2, 3, 4],
                           2: [2, 3, 4, 1],
                           3: [3, 4, 1, 2],
                           4: [4, 1, 2, 3]}

def test_1_plurality():
    expected_output = 4
    print("Expected Output:", expected_output)
    actual_output = voting.plurality(plurality_test_1_values, 4)
    print("Actual Output:", actual_output)
    
# test_1_plurality()

veto_test_1_values = {1: [1, 2, 3, 4],
                      2: [2, 3, 4, 1],
                      3: [3, 4, 1, 2],
                      4: [4, 1, 2, 3]}

def test_1_veto():
    expected_output = 4
    print("Expected Output:", expected_output)
    actual_output = voting.veto(veto_test_1_values, 4)
    print("Actual Output:", actual_output)
    
# test_1_veto()

borda_test_1_values = {1: [1, 2, 3, 4],
                      2: [2, 3, 4, 1],
                      3: [3, 4, 1, 2],
                      4: [4, 1, 2, 3]}

def test_1_borda():
    expected_output = 4
    print("Expected Output:", expected_output)
    actual_output = voting.borda(borda_test_1_values, 4)
    print("Actual Output:", actual_output)
    
# test_1_borda()

harmonic_test_1_values = {1: [1, 2, 3, 4],
                      2: [2, 3, 4, 1],
                      3: [3, 4, 1, 2],
                      4: [4, 1, 2, 3]}

def test_1_harmonic():
    expected_output = 2
    print("Expected Output:", expected_output)
    actual_output = voting.harmonic(harmonic_test_1_values, "min")
    print("Actual Output:", actual_output)
    
# test_1_harmonic()
    
def test_range_voting():
    expected_output = 4
    print("Expected Output:", expected_output)
    actual_output = voting.range_voting(values, 7)
    print("Actual Output:", actual_output)
    
# test_range_voting()