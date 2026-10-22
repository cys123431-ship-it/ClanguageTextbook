#include <stdio.h>

long sum_proper_divisors(int number)
{
    if (number < 2 || number > 1000) {
        return -1L;
    }
    long sum = 0L;
    for (int divisor = 1; divisor <= number / 2; ++divisor) {
        if (number % divisor == 0) {
            sum += divisor;
        }
    }
    return sum;
}

int is_perfect(int number)
{
    long sum = sum_proper_divisors(number);
    if (sum < 0L) {
        return -1;
    }
    return sum == number;
}

void report(int number)
{
    long sum = sum_proper_divisors(number);
    int perfect = is_perfect(number);
    if (sum < 0L) {
        printf("n=%d: invalid\n", number);
    } else {
        printf("n=%d sum=%ld perfect=%d\n", number, sum, perfect);
    }
}

int main(void)
{
    report(6);
    report(28);
    report(12);
    report(1);
    return 0;
}
