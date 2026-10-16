#include <stdio.h>

int clamp_score(int score)
{
    if (score < 0)
        return 0;
    if (score > 100)
        return 100;
    return score;
}

int pass_count(int a, int b, int c)
{
    int passed = 0;
    if (clamp_score(a) >= 60)
        passed++;
    if (clamp_score(b) >= 60)
        passed++;
    if (clamp_score(c) >= 60)
        passed++;
    return passed;
}

int main(void)
{
    printf("Passed: %d\n", pass_count(58, 60, 140));
    return 0;
}
