def get_valid_input(prompt, min_value, max_value):
    """
    Get a valid float input from the user within the specified range.
    """
    while True:
        try:
            value = float(input(prompt))
            if min_value <= value <= max_value:
                return value
            else:
                print(f"Please enter a value between {min_value} and {max_value}.")
        except ValueError:
            print("Invalid input. Please enter a number.")

def calculate_priority_scores(tasks):
    """
    Calculate the priority scores of tasks based on user-defined weights.

    :param tasks: A dictionary of tasks with their urgency and importance values.
    :return: List of tuples (task, priority score) sorted by priority score in descending order.
    """
    # Get user input for weight factors
    print("Please enter the weight factors for urgency and importance. They should sum up to 1.")
    a = get_valid_input("Enter weight factor for urgency (a): ", 0, 1)
    b = get_valid_input("Enter weight factor for importance (b): ", 0, 1)

    # Validate the sum of a and b
    while a + b != 1:
        print("The sum of urgency and importance factors should be 1. Please re-enter the values.")
        a = get_valid_input("Enter weight factor for urgency (a): ", 0, 1)
        b = get_valid_input("Enter weight factor for importance (b): ", 0, 1)

    # Calculate priority scores
    priority_scores = {task: (urgency * a + importance * b) for task, (urgency, importance) in tasks.items()}
    return sorted(priority_scores.items(), key=lambda x: x[1], reverse=True)

# Example usage
tasks = {
    "Email client about project": (9, 8),
    "Prepare presentation for next week": (4, 7),
    # ... other tasks ...
}

priority_scores_sorted = calculate_priority_scores(tasks)
for task, score in priority_scores_sorted:
    print(f"{task}: Priority Score = {score}")
