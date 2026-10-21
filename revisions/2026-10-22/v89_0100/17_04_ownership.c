/* C11: an owning pointer and a borrowed pointer have different responsibilities. */
#include <stdio.h>
#include <stdlib.h>

static int make_scores(int **out)
{
    if (out == NULL || *out != NULL) {
        return 0;
    }
    int *temporary = malloc(4 * sizeof *temporary);
    if (temporary == NULL) {
        return 0;
    }
    temporary[0] = 70;
    temporary[1] = 80;
    temporary[2] = 90;
    temporary[3] = 60;
    *out = temporary;
    return 1;
}

static int sum_scores(const int *borrowed, size_t count, int *out)
{
    if (borrowed == NULL || out == NULL || count != 4) {
        return 0;
    }
    int total = 0;
    for (size_t i = 0; i < count; ++i) {
        total += borrowed[i];
    }
    *out = total;
    return 1;
}

static void release_scores(int **owner)
{
    if (owner != NULL) {
        free(*owner);
        *owner = NULL;
    }
}

int main(void)
{
    int *owned = NULL;
    if (!make_scores(&owned)) {
        fputs("allocation failed\n", stderr);
        return EXIT_FAILURE;
    }
    const int *borrowed = owned;
    int total = 0;
    if (!sum_scores(borrowed, 4, &total)) {
        release_scores(&owned);
        return EXIT_FAILURE;
    }
    printf("count=4 sum=%d\n", total);
    release_scores(&owned);
    borrowed = NULL;
    printf("owner_cleared=%s\n", owned == NULL ? "yes" : "no");
    release_scores(&owned);
    puts("second_cleanup=safe");
    return EXIT_SUCCESS;
}
