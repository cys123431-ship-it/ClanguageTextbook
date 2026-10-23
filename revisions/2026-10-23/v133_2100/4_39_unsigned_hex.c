#include <stdio.h>
#include <limits.h>

int main(void)
{
    int negative = -1;
    unsigned int converted = (unsigned int)negative;
    unsigned int minus_two = (unsigned int)-2;
    unsigned int minus_three = (unsigned int)-3;
    printf("Octal 255: %o\n", 255U);
    printf("Hex 255: %x\n", 255U);
    printf("Minus two as hex: %x\n", minus_two);
    printf("Minus three as hex: %x\n", minus_three);
    printf("Signed: %d\n", negative);
    printf("Unsigned: %u\n", converted);
    printf("Hex: %x\n", converted);
    printf("UINT_MAX: %u\n", UINT_MAX);
    return 0;
}
