#include <limits.h>
#include <stdio.h>

static void print_digits(unsigned int value)
{
    if (value >= 10u) {
        print_digits(value / 10u);
    }
    putchar('0' + (int)(value % 10u));
}

static void print_decimal(int value)
{
    if (value < 0) {
        putchar('-');
        print_digits(0u - (unsigned int)value);
    } else {
        print_digits((unsigned int)value);
    }
}

int main(void)
{
    print_decimal(-1234);
    putchar('\n');
    print_decimal(0);
    putchar('\n');
    print_decimal(405);
    putchar('\n');
    print_decimal(INT_MIN);
    putchar('\n');
    return 0;
}
