def generate_preferences(values):
    """
    Gets the preferences from the worksheet, outputs a preference profile.

    Collects the agent's preferences from the worksheet. It then returns a dictionary
    with each agent and the order of their preferences being ranked from 1 (the most preferrred)
    to the least preferred. The order of the ranked preferences reflects the original
    positions of the agent's preferences in the worksheet.

    Parameters:
        values (openpyxl worksheet): the numerical values each agent assigns to the alternatives
        (a higher value means an alternative is more preferred).

    Returns:
        preferences (dict): a dictionary of agents and preferences. An example entry would be 1: [2, 4, 1, 3]
        where the key '1:' is Agent 1 and the values [2, 4, 1, 3] are their preferred alternatives ranked
        where the order in that list of values is congruous with the order of the values in the worksheet.
    """
    preferences = {}
    # First agent is 1. Loops through the worksheet's rows to create the dictionary in the needed format.
    agent = 1
    for row in values.iter_rows(values_only=True):
        # Makes each row of agent's valuations into a list.
        valuation_list = list(row)
        preferences_list = []
        valuation_list_dictionary = {}
        # Builds the valuation_list_dictionary where the key is the valuation of an alternative
        # and the values are their indices in the valuation_list, it is then sorted by valuation from largest to smallest.
        for index, valuation in enumerate(valuation_list):
            valuation_list_dictionary.setdefault(valuation, []).append(index + 1)
        valuation_list_dictionary_sorted = dict(sorted(valuation_list_dictionary.items(), reverse=True))
        # Goess through each list in each value in the dictionary and sorts them largest to smallest.
        # This accounts for duplicate valuations. Then it constructs the list of preferences.
        for agent_valuation_list in valuation_list_dictionary_sorted:
            agent_valuation_list_sorted = valuation_list_dictionary_sorted[agent_valuation_list]
            agent_valuation_list_sorted.sort(reverse=True)
            # This append to preferences_list the sorted lists which are the values in the dictionary.
            preferences_list = preferences_list + valuation_list_dictionary_sorted[agent_valuation_list]
        # Adds the agent and their preferences_list to the preferences dictionary.
        preferences[agent] = preferences_list
        agent = agent + 1
    return preferences

def points_tally(preferences):
    """
    This tallies the points for each alternative where a voting rule inputs a preference profile.

    For each voting rule function, a points dictionary is necessary.
    This dictionary tallies the points for each alternative to provide a way to identify the winners for each voting rule.
    This works for each voting rule where a preference profile (preferences) is inputted.
    Thus, the range_voting function which does not input a preference_profile uses an alternative points tally within its function.

    Parameters:
        preferences (dict): a dictionary of agents and preferences.

    Returns:
        points (dict): a dictionary where alternatives are the keys and the points are the values.
    """
    points = {}
    alternative_key = 1
    while alternative_key <= len(preferences[1]):
        points[alternative_key] = 0
        alternative_key = alternative_key + 1
    return points

def tie_checker(points):
    """
    A function to check if a tie has occured.

    Checks the alternatives with the highest scores for a voting rule function.
    Identifies if there is more than one alternative with a high score, a.k.a. if there is a tie or not.

    Parameters:
        points (dict): A dictionary where the keys are the alternatives and the values are their scores.

    Returns:
        high_scores_list (list): The alternatives with the highest score.
    """
    highest_score = max(list(points.values()))
    high_scores_list = []
    for alternative in points:
        if points[alternative] == highest_score:
            high_scores_list.append(alternative)
        else:
            pass
    return high_scores_list

def tie_breaker(preferences, tie_break, high_scores_list):
    """
    A function for executing the tie-break rules.

    A tie-break is decided by either: the "max" rule where the alternative with the highest number is the winner.
    The "min" rule where the alternative with the lowest number is the winner.
    The agent i rule where the alternative which is ranked highest in agent i's preferences is the winner.

    Parameters:
        preferences (dict/openpyxl worksheet): either the preferences dictionary from the generate_preferences function
        or for the range_voting function the values from the openpyxl worksheet.
        tie_break (str or int): either "max", "min" or an integer i. If i does not correspond with an agent an error is raised.
        high_scores_list (list): a list of tied alternatives for a voting rule function.

    Returns:
        winner (int): the winning alternative for a voting rule when a tie-break is necessary.
    """
    try:
        if isinstance(tie_break, int) is True:
            for agent_i_alternative in preferences[tie_break]:
                if agent_i_alternative in high_scores_list:
                    winner = agent_i_alternative
                    break
        elif tie_break == "min":
            winner = min(high_scores_list)
        elif tie_break == "max":
            winner = max(high_scores_list)
        else:
            print("""Please enter "min", "max" or an integer.""")
    except KeyError:
        print(f"Agent {tie_break} does not exist.")
    return winner

def dictatorship(preferences, agent):
    """
    An agent is selected, the winner is the one that agent ranks first.

    Parameters:
        preferences (dict): the preferences returned from the generate_preferences() function.
        agent (int): the number of the agent. An error message is printed if the number does not correspond to an agent.

    Returns:
        winner (int): the winning alternative.
    """
    if agent in preferences:
        # Iterates through the dictionary to find the agent, returns their first choice.
        for agent_profile in preferences:
            if agent_profile == agent:
                return int(preferences[agent_profile][0])
            else:
                pass
    else:
        raise ValueError(f"Inputted integer {agent} does not correspond to an agent.")

def scoring_rule(preferences, score_vector, tie_break):
    """
    The winner is the alternative with the highest score.

    The score_vector gives a score to each alternative in each agent's preferences.
    For each agent, the highest score is given to the top ranked alternative and so on.
    The score_vector must be the same number of scores as the number of alternatives, or an error is given.
    It then sums the scores for each alternative and returns the winner.
    A tie-breaking rule is used in the event of a draw.

    Parameters:
        preferences (dict): the preferences returned from the generate_preferences() function.
        score_vector (list of floats): the scores to be given to the alternatives.
        tie_break (str or int): either "max", "min" or an integer i. If i does not correspond with an agent an error is raised.

    Returns:
        winner (int): the winning alternative.
        or
        False: where the score_vector does not have the same number of scores as the number of alternatives.
    """
    # Checks if the score vector's length is the same as the number of alternatives.
    if len(preferences[1]) == len(score_vector):
        # Populates the keys (alternatives) in the points dictionary.
        points = points_tally(preferences)
        # Populates the values (total scores) in the points dictionary.
        for agent_profile in preferences:
            preference_score_list = zip(preferences[agent_profile], sorted(score_vector, reverse=True))
            for preference_tuple_value, score_tuple_value in list(preference_score_list):
                points[preference_tuple_value] += score_tuple_value
        # Finds the high score, then the alternative(s) with that score.
        high_scores_list = tie_checker(points)
        # Identifies if there is a tie, if so calls the tie_breaker function.
        if len(high_scores_list) > 1:
            winner = tie_breaker(preferences, tie_break, high_scores_list)
        else:
            winner = high_scores_list[0]
        return int(winner)
    else:
        print("Incorrect input")
        return False

def plurality(preferences, tie_break):
    """
    The winner is the alternative which appears the most in the first positions of agents' preferences.

    Finds the alternative which is ranked first for each agent.
    Then identifies which alternative(s) appear the most out of those.
    A tie-breaking rule is used in the event of a draw.

    Parameters:
        preferences (dict): the preferences returned from the generate_preferences() function.
        tie_break (str or int): either "max", "min" or an integer i. If i does not correspond with an agent an exception is raised.

    Returns:
        winner (int): the winning alternative.
    """
    # Populates the keys (alternatives) in the points dictionary.
    points = points_tally(preferences)
    # Gets the first ranked alternative of each agent.
    ranked_first = []
    for agent in preferences:
        ranked_first.append(preferences[agent][0])
        # Finds how many times each alternative was an agent's first rank.
        ranked_first_preference = 0
        for ranked_first_preference in ranked_first:
            points[preferences[agent][0]] = ranked_first.count(ranked_first_preference)
    # Finds the high score, then the alternative(s) with that score.
    high_scores_list = tie_checker(points)
    # Identifies if there is a tie, if so calls the tie_breaker function.
    if len(high_scores_list) > 1:
        winner = tie_breaker(preferences, tie_break, high_scores_list)
    else:
        winner = high_scores_list[0]
    return int(winner)

def veto(preferences, tie_break):
    """
    The winner is the alternative with the most points, where every alternative apart from the last ranked one gets 1 point.

    Each alternative in an agent's preferences receives 1 point, apart from each last ranked alternative.
    The alternative with the most points is the winner. A tie-breaking rule is used in the event of a draw.

    Parameters:
        preferences (dict): the preferences returned from the generate_preferences() function.
        tie_break (str or int): either "max", "min" or an integer i. If i does not correspond with an agent an exception is raised.

    Returns:
        winner (int): the winning alternative.
    """
    # Populates the keys (alternatives) in the points dictionary.
    points = points_tally(preferences)
    # Populates the values (total scores) in the points dictionary.
    for agent in preferences:
        preferences_list = preferences[agent]
        for alternative in preferences_list[0:-1]:
            points[alternative] += 1
    # Finds the high score, then the alternative(s) with that score.
    high_scores_list = tie_checker(points)
    # Identifies if there is a tie, if so calls the tie_breaker function.
    if len(high_scores_list) > 1:
        winner = tie_breaker(preferences, tie_break, high_scores_list)
    else:
        winner = high_scores_list[0]
    return int(winner)

def borda(preferences, tie_break):
    """
    The winner is the alternative with the most points according to the borda voting rule.

    An agent's alternative at position j receives a score of m - j where m is the number of alternatives.
    The alternative with the highest score is the winner. A tie-breaking rule is used in the event of a draw.

    Parameters:
        preferences (dict): the preferences returned from the generate_preferences() function.
        tie_break (str or int): either "max", "min" or an integer i. If i does not correspond with an agent an exception is raised.

    Returns:
        winner (int): the winning alternative.
    """
    # Populates the keys (alternatives) in the points dictionary.
    points = points_tally(preferences)
    # Populates the values (total scores) in the points dictionary.
    for agent in preferences:
        preferences_list = preferences[agent]
        for alternative in preferences_list:
            score = len(preferences_list) - preferences_list.index(alternative) - 1
            points[alternative] += score
    # Finds the high score, then the alternative(s) with that score.
    high_scores_list = tie_checker(points)
    # Identifies if there is a tie, if so calls the tie_breaker function.
    if len(high_scores_list) > 1:
        winner = tie_breaker(preferences, tie_break, high_scores_list)
    else:
        winner = high_scores_list[0]
    return int(winner)

def harmonic(preferences, tie_break):
    """
    The winner is the alternative with the most points according to the harmonic voting rule.

    Every agent's alternative receives a score where the alternative at position j has a score of 1/j.
    The alternative with the highest score is the winner. A tie-breaking rule is used in the event of a draw.

    Parameters:
        preferences (dict): the preferences returned from the generate_preferences() function.
        tie_break (str or int): either "max", "min" or an integer i. If i does not correspond with an agent an exception is raised.

    Returns:
        winner (int): the winning alternative.
    """
    # Populates the keys (alternatives) in the points dictionary.
    points = points_tally(preferences)
    # Populates the values (total scores) in the points dictionary.
    for agent in preferences:
        preferences_list = preferences[agent]
        for alternative in preferences_list:
            score = 1/(preferences_list.index(alternative) + 1)
            points[alternative] += score
    # Finds the high score, then the alternative(s) with that score.
    high_scores_list = tie_checker(points)
    # Identifies if there is a tie, if so calls the tie_breaker function.
    if len(high_scores_list) > 1:
        winner = tie_breaker(preferences, tie_break, high_scores_list)
    else:
        winner = high_scores_list[0]
    return int(winner)

def STV(preferences, tie_break):
    """
    The winner is the last alternative remaining.

    There are a number of rounds. Each round the alternative which appears least frequently in an agent's first place is eliminated.
    When an alternative is eliminated the remaining alternatives move across a position in an agent's rankings to take that space.
    The last remaining alternative is the winner.
    A tie-breaking rule is used in the event of a draw.

    Parameters:
        preferences (dict): the preferences returned from the generate_preferences() function.
        tie_break (str or int): either "max", "min" or an integer i. If i does not correspond with an agent an exception is raised.

    Returns:
        stv_winner (int): the winning alternative.
    """
    # Populates the keys (alternatives) in the points dictionary.
    points = points_tally(preferences)
    # A while loop for going through the preferences dicitonary multiple times.
    while True:
        # This makes column_list which is the first alternative from each agent's
        # list of alternatives in the preferences dictionary.
        for alternative in sorted(list(preferences[1])):
            column_list = []
            for agents_alternatives_list in preferences.values():
                column_list.append(agents_alternatives_list[alternative-1])
            break
        # The count of each alternative in the current column list is put into the points dictionary.
        for alternative in column_list:
            points[alternative] += 1
        # This takes the alternative(s) with the lowest number of appearances in the agent's first place
        # and puts them into a list called alternatives_appearances. i.e. those alternatives with the least points in the points dictionary.
        alternatives_appearances = []
        min_alternative_appearance = min(points.values())
        for alternative in points:
            if points[alternative] == min_alternative_appearance:
                alternatives_appearances.append(alternative)
            else:
                pass
        # This removes the alternative(s) in alternatives_appearances from the preferences dictionary.
        if len(alternatives_appearances) == len(preferences[1]):
            pass
        else:
            for alternative in alternatives_appearances:
                for agents_alternatives_list in preferences.values():
                    for original_alternative in agents_alternatives_list:
                        if original_alternative == alternative:
                            agents_alternatives_list.remove(alternative)
                        else:
                            pass
                # Removes those alternatives from the points dictionary as well.
                points.pop(alternative)
        # This piece of code checks how many points the remaining alternatives have in the points dictionary.
        # It is then converted to a set to remove duplicates. This code prevents the situation where all
        # alternatives may be removed in the event of the last remaining alternatives all having the lowest points.
        stv_alternative_list = []
        for stv_alternative in points.values():
            stv_alternative_list.append(stv_alternative)
        stv_alternative_set = set(stv_alternative_list)
        # Finds the high score, then the alternative(s) with that score.
        high_scores_list = tie_checker(points)
        # This asks if there is only one element in the set. If there is then the while loop can now finish.
        if len(stv_alternative_set) == 1:
            # Identifies if there is a tie, if so calls the tie_breaker function.
            if len(high_scores_list) > 1:
                winner = tie_breaker(preferences, tie_break, high_scores_list)
            elif len(high_scores_list) == 1:
                winner = preferences[1][0]
            return int(winner)
        else:
            # Clears the dictionary values after each loop of the while loop.
            points = points.fromkeys(points, 0)
            continue

def range_voting(values, tie_break):
    """
    The winner is the alternative with the largest sum of values.

    The values for each alternative given by the agents in the .xlsx file
    are summed, the winner is the alternative with the largest sum of values.
    A tie-breaking rule is used in the event of a draw.

    Parameters:
        values (openpyxl worksheet): the numerical values each agent assigns to the alternatives
        (a higher value means an alternative is more preferred).
        tie_break (str or int): either "max", "min" or an integer i. If i does not correspond with an agent an exception is raised.

    Returns:
        winner (int): the winning alternative.
    """
    preferences = generate_preferences(values)
    # Finds the row length to help construct the points dictionary.
    for values_row in values.iter_rows(values_only=True):
        row_length = len(list(values_row))
    points = {}
    # Populates the keys (alternatives) in the points dictionary.
    alternative_key = 1
    while alternative_key <= row_length:
        points[alternative_key] = []
        alternative_key = alternative_key + 1
    # This arranges the values by alternatives in the points dictionary.
    # e.g. Alternative 1: [Agent 1's 1st preference... Agent n's first preference]
    for values_row in values.iter_rows(values_only=True):
        loop_iterator = 0
        while loop_iterator < row_length:
            points[loop_iterator + 1].append(values_row[loop_iterator])
            loop_iterator = loop_iterator + 1
    # Sums the totals for each alternative in the points dictionary.
    for alternative in points:
        points[alternative] = sum(points[alternative])
    # Finds the high score, then the alternative(s) with that score.
    high_scores_list = tie_checker(points)
    # Identifies if there is a tie, if so calls the tie_breaker function.
    if len(high_scores_list) > 1:
        winner = tie_breaker(preferences, tie_break, high_scores_list)
    else:
        winner = high_scores_list[0]
    return int(winner)
