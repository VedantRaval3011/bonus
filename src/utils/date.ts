export function calculateServiceMonths(doj: Date, currentDate: Date): number {
  const diffTime = currentDate.getTime() - doj.getTime();
  const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
  return Math.floor(diffDays / 30.44); // Average days per month
}

export function formatDate(date: Date): string {
  return date.toLocaleDateString('en-IN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric'
  });
}
