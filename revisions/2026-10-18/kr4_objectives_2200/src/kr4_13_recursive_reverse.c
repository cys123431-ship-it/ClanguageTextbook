#include <stdio.h>
#include <string.h>

static void reverse_range(char text[], size_t left, size_t right)
{
    if (left >= right) {
        return;
    }
    char temp = text[left];
    text[left] = text[right];
    text[right] = temp;
    reverse_range(text, left + 1u, right - 1u);
}

static void reverse_text(char text[])
{
    size_t length = strlen(text);
    if (length > 0u) {
        reverse_range(text, 0u, length - 1u);
    }
}

int main(void)
{
    char first[] = "algorithm";
    char second[] = "ABCD";
    char empty[] = "";
    reverse_text(first);
    reverse_text(second);
    reverse_text(empty);
    printf("first=%s\nsecond=%s\nempty=[%s]\n", first, second, empty);
    return 0;
}
