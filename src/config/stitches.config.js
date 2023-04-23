import { createStitches } from '@stitches/react'

export const { config, createTheme, css, getCssText, globalCss, styled, theme, keyframes } =
  createStitches({
    theme: {
      colors: {
        primary1: '#006BFF',
        primary2: '#002646',
        primary3: '#30D2FF',
        primary4: '#E2F2FF',

        typographyHigh: 'rgba(0, 0, 0, 0.87)',
        typographyMedium: 'rgba(0, 0, 0, 0.6)',
        typographyLow: 'rgba(0, 0, 0, 0.38)',

        neutral1: '#FFFFFF',
        neutral2: '#FBFBFB',
        neutral3: '#F1F1F1',
        neutral4: '#E3E3E3',
        neutral5: '#CECECE',
        neutral6: '#B7B7B7',
        neutral7: '#949494',
        neutral8: '#7C7C7C',
        neutral9: '#646464',
        neutral10: '#4B4B4B',
        neutral11: '#3C3C3C',
        neutral12: '#2C2C2C',
        neutral13: '#202020',
        neutral14: '#0C0C0C',
        neutral15: '#000000',

        gray100: 'hsl(206,22%,99%)',
        gray200: 'hsl(206,12%,97%)',
        gray300: 'hsl(206,11%,92%)',
        gray400: 'hsl(206,10%,84%)',
        gray500: 'hsl(206,10%,76%)',
        gray600: 'hsl(206,10%,44%)',

        purple100: 'hsl(252,100%,99%)',
        purple200: 'hsl(252,100%,98%)',
        purple300: 'hsl(252,100%,94%)',
        purple400: 'hsl(252,75%,84%)',
        purple500: 'hsl(252,78%,60%)',
        purple600: 'hsl(252,80%,53%)',
      },
      space: {
        1: '5px',
        2: '10px',
        3: '15px',
        4: '20px',
        5: '25px',
        6: '35px',
      },
      sizes: {
        1: '5px',
        2: '10px',
        3: '15px',
        4: '20px',
        5: '25px',
        6: '35px',
      },
      fontSizes: {
        1: '12px',
        2: '13px',
        3: '15px',
        4: '17px',
        5: '19px',
        6: '21px',
      },
      fonts: {
        system: 'system-ui',
      },
    },
    utils: {
      marginX: (value) => ({
        marginLeft: value,
        marginRight: value,
      }),
      marginY: (value) => ({
        marginTop: value,
        marginBottom: value,
      }),
      paddingX: (value) => ({
        paddingLeft: value,
        paddingRight: value,
      }),
      paddingY: (value) => ({
        paddingTop: value,
        paddingBottom: value,
      }),
      alignCenter: () => ({
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
      }),
    },
    media: {
      bp1: '(max-width: 1200px)',
      bp2: '(max-width: 992px)',
      bp3: '(max-width: 768px)',
      bp4: '(max-width: 576px)',
      bp5: '(max-width: 1020px)',
    },
  })
