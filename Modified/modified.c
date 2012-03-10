/************************************************************/
/* TTDX Editor - MODIFIED.EXE Application                   */
/* Version 1.10.18                                          */
/*                                                          */
/* Copyright (c) Owen Rudge 2003-2004. All Rights Reserved. */
/************************************************************/
/* Revision History:
 *
 * 13/09/2003: Updated to work with TTDPatch 2.0 beta
 *             5.1 MG or higher (orudge)
 *
 * 19/06/2004: Updated to add support for listing modifications (orudge)
 */
 
#include <stdio.h>
#include <string.h>

char version_greater(char ma1, char mi1, char r1, char ma2, char mi2, char r2);
void list_modifications(char mod);

int main(int argc, char *argv[])
{
    FILE *in;
    char major, minor, revision, modifications;
    long newformat;
    
    if (argc != 2)
    {
        if (argc == 1)
        {
           printf("Usage: %s filename\n", argv[0]);
           printf("\n");
           printf("'filename' should be an uncompressed ('big') Transport Tycoon Deluxe\n");
           printf("saved game or scenario (eg, created by 'Save Uncompressed' in TTDX\n");
           printf("Editor, or decoded with SV1Codec)\n");
           return(1);
        }
        else if (argc == 3)
        {
            if (stricmp(argv[1], "/r") == 0)
            {
                in = fopen(argv[2], "wb");

                if (in == NULL)
                {
                    printf("Error: Unable to open %s for writing!\n", argv[1]);
                    return(1);
                }

                if (fseek(in, 0x44CB4, SEEK_SET) != 0)
                {
                    printf("Error: Unable to seek to 0x44CB4. File may not be a decompressed TTD saved game.\n");
                    fclose(in);
                    return(1);
                }

                fread(&newformat, 4, 1, in);

                if (newformat != 0x70445454)
                {
                    if (fseek(in, 0x24CCB, SEEK_SET) != 0)
                    {
                        printf("Error: Unable to seek to 0x24CCB. File may not be a decompressed TTD saved game.\n");
                        fclose(in);
                        return(1);
                    }
    
                    fputc(0, in);
                    fputc(0, in);
                    fputc(0, in);

                    printf("This saved game has had its modification signature reset.\n");
                    return(0);
                }
                else
                {
                    if (fseek(in, 0x44BBA, SEEK_SET) != 0)
                    {
                        printf("Error: Unable to seek to 0x44BBA. File may not be a decompressed TTD saved game.\n");
                        fclose(in);
                        return(1);
                    }
    
                    fputc(0, in);
                    fputc(0, in);
                    fputc(0, in);

                    printf("This saved game has had its modification signature reset.\n");
                    return(0);
                }
            }                                   
            else
            {
               printf("Usage: %s filename\n", argv[0]);
               printf("\n");
               printf("'filename' should be an uncompressed ('big') Transport Tycoon Deluxe\n");
               printf("saved game or scenario (eg, created by 'Save Uncompressed' in TTDX\n");
               printf("Editor, or decoded with SV1Codec)\n");
               return(1);
            }
        }
    }
    else
    {
        in = fopen(argv[1], "rb");

        if (in == NULL)
        {
            printf("Error: Unable to open %s for reading!\n", argv[1]);
            return(1);
        }

        if (fseek(in, 0x44CB4, SEEK_SET) != 0)
        {
            printf("Error: Unable to seek to 0x44CB4. File may not be a decompressed TTD saved game.\n");
            fclose(in);
            return(1);
        }

        fread(&newformat, 4, 1, in);
        
        if (newformat != 0x70445454)
        {
            if (fseek(in, 0x24CCB, SEEK_SET) != 0)
            {
                printf("Error: Unable to seek to 0x24CCB. File may not be a decompressed TTD saved game.\n");
                fclose(in);
                return(1);
            }

            major = fgetc(in);
            minor = fgetc(in);
            revision = fgetc(in);
            modifications = fgetc(in);
            
            if (major == 0 && minor == 0 && revision == 0)
                printf("This saved game has not been modified by TTDX Editor 1.10.0014 or later.\n");
            else
            {
                printf("This saved game has been modified by TTDX Editor %d.%02d.%04d.\n", major, minor, revision);

                if (version_greater(major, minor, revision, 1, 10, 18))
                   list_modifications(modifications);
            }
        }
        else
        {
            if (fseek(in, 0x44BBA, SEEK_SET) != 0)
            {
                printf("Error: Unable to seek to 0x44BBA. File may not be a decompressed TTD saved game.\n");
                fclose(in);
                return(1);
            }

            major = fgetc(in);
            minor = fgetc(in);
            revision = fgetc(in);
            modifications = fgetc(in);

            if (major == 0 && minor == 0 && revision == 0)
                printf("This saved game has not been modified by TTDX Editor 1.10.0014 or later.\n");
            else
            {
                printf("This saved game has been modified by TTDX Editor %d.%02d.%04d.\n", major, minor, revision);

                if (version_greater(major, minor, revision, 1, 10, 18))
                   list_modifications(modifications);
            }
        }

        fclose(in);
    }

    return(0);
}

char version_greater(char ma1, char mi1, char r1, char ma2, char mi2, char r2)
{
    if (ma1 > ma2)
       return(1);

    if (ma1 == ma2)
    {
        if (mi1 > mi2)
           return(1);

        if (mi1 == mi2)
        {
            if (r1 > r2)
               return(1);

            if (r1 == r2)
               return(1);
        }
    }

    return(0);
}

void list_modifications(char mod)
{
    puts("\nThe following modifications have been made to the saved game:");
    
    if (mod & 1)
       puts("- Player data updated");

    if (mod & 2)
       puts("- City data updated");
    
    if (mod & 4)
       puts("- Industry data updated");

    if (mod & 8)
       puts("- Stations data updated");

    if (mod & 16)
       puts("- Vehicles data updated");

    if (mod & 32)
       puts("- Terrain data updated");

    if (mod & 64)
       puts("- Other unknown data updated");

    if (mod & 128)
       puts("- Other unknown data updated");
}
