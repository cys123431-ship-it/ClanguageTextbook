#include <stdio.h>

int main(void)
{
    int code = 0;

    printf("Printable ASCII code (32 to 126): ");
    if (scanf("%d", &code) != 1 || code < 32 || code > 126) {
        puts("Invalid ASCII code.");
        return 1;
    } else {
        if ('A' != 65 || 'a' != 97 || '0' != 48) {
            puts("This example requires ASCII-compatible characters.");
            return 2;
        } else {
            printf("Character: %c\n", (char)code);
        }
    }
    return 0;
}
