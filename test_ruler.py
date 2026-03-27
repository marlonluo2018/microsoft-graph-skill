# Test time ruler generation
time_labels = " | "
for hour in range(24):
    if hour < 10:
        time_labels += f"│{hour}"
    else:
        time_labels += f"{hour}"

print(time_labels)
print(f"Length: {len(time_labels)}")
print(f"Expected: {3 + 48} (3 for ' | ' + 48 for 24 hours * 2 chars)")

# Print each character with its position
for i, char in enumerate(time_labels):
    print(f"{i}: '{char}' (ord={ord(char)})")

# Made with Bob
