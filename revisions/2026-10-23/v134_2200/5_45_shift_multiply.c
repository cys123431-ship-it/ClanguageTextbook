#include <stdio.h>
int main(void)
{
    unsigned int value = 0U;
    unsigned int shift = 0U;
    printf("Value (0..1000) and shift (0..4): ");
    if (scanf("%u %u", &value, &shift) != 2 ||
        value > 1000U || shift > 4U) {
        puts("Invalid input.");
        return 1;
    } else {
        unsigned int shifted = value << shift;
        printf("Shifted: %u\n", shifted);
        printf("Restored: %u\n", shifted >> shift);
    }
    return 0;
}