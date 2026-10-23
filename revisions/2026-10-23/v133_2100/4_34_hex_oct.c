#include <stdio.h>

int main(void)
{
    unsigned int value = 0U;
    printf("Enter a hexadecimal number: ");
    if (scanf("%x", &value) != 1) {
        puts("Invalid hexadecimal input.");
        return 1;
    } else {
        printf("Octal: %o\n", value);
        printf("Decimal: %u\n", value);
        printf("Hexadecimal: %#x\n", value);
    }
    return 0;
}
