export function hasNumber(s: string): boolean {
  return /\d/.test(s);
}

export function separateHumanizedTimeForTranslating(
  input: string,
): [string, number] {
  // Extract the number
  const number = parseInt(input.replace(/\D/g, ''), 10);

  // Replace the number with '%d'
  const text = input.replace(/[0-9]+/, '%d');

  return [text, number];
}
