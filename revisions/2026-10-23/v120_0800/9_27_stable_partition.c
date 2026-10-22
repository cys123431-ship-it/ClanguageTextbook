#include <stddef.h>
#include <stdio.h>

int main(void)
{
    int data[] = {7, 2, 9, 4, 6, 3, 8, 1};
    size_t count = sizeof data / sizeof data[0];
    int result[sizeof data / sizeof data[0]];
    size_t used = 0;

    for (size_t i = 0; i < count; ++i) {
        if (data[i] % 2 == 0) {
            result[used] = data[i];
            ++used;
        }
    }
    for (size_t i = 0; i < count; ++i) {
        if (data[i] % 2 != 0) {
            result[used] = data[i];
            ++used;
        }
    }
    for (size_t i = 0; i < count; ++i) {
        printf("%s%d", i == 0 ? "" : " ", result[i]);
    }
    putchar('\n');
    return 0;
}