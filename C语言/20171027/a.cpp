#define _XOPEN_SOURCE  
//#include <unistd.h>  
#include <fcntl.h>  
#include <stdio.h>  
#include <stdlib.h>  
#include <pty.h>  
  
// pty master  
#define PTMASTER "/dev/ptmx"  
  
int main()  
{  
  
        int amaster, aslave;  
        char *slavename;  
        int masterfd;  
  
        masterfd = openpty(&amaster, &aslave, NULL, NULL, NULL);  
        slavename = ptsname(amaster);  
        printf("pts name : %s\n", slavename);  
  
        masterfd = open(PTMASTER, O_RDWR);  
        if (masterfd < 0) {  
                perror("open");  
                exit(EXIT_FAILURE);  
        }  
        slavename = ptsname(masterfd);  
        if (slavename == NULL) {  
                printf ("Get pts name failed\n");  
                exit (EXIT_FAILURE);  
        }  
        printf ("pts name : %s\n", slavename);  
        close(masterfd);  
  
        return 0;  
}  