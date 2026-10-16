#include <stdio.h>
#include <limits.h>
#include <stdint.h>

int main(void)
{
    int_least32_t points = 250;
    printf("byte_bits=%d\n", CHAR_BIT);
    printf("least32_bytes=%zu\n", sizeof(points));
    printf("points_positive=%d\n", points > 0);
    return 0;
}
