import json

def load_json(file_path):
	with open(file_path, 'r') as file:
		return json.load(file)

def find_invalid_connectors(data):
	shapes = {shape['id'] for shape in data.get('shapes', [])}
	invalid_connectors = []

	for connector in data.get('connectors', []):
		from_id = connector.get('from_id')
		to_id = connector.get('to_id')
		if from_id not in shapes or to_id not in shapes:
			invalid_connectors.append(connector)

	return invalid_connectors

def main():
	input_file_path = r"C:\Users\leroy\Downloads\nwidd\CAD_interactions_report.json"
	data = load_json(input_file_path)
	invalid_connectors = find_invalid_connectors(data)

	if invalid_connectors:
		print("Invalid Connectors:")
		for connector in invalid_connectors:
			print(connector)
	else:
		print("All connectors are valid.")

if __name__ == "__main__":
	main()