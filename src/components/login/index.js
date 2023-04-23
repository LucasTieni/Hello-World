import { styled } from 'config/stitches.config'

export const Container = styled('div', {
  width: '350px',
  height: '297px',
  paddingTop: '40px',
  margin: '0 auto',
})

export const Button = styled('button', {
  width: '100%',
  borderRadius: 4,
  border: 'none',
  padding: '16px',
  background: '$primary1',
  color: '#fff',
  fontSize: 18,
  lineHeight: 1,
  fontWeight: 500,
  cursor: 'pointer',
})

export const Input = styled('input', {
  width: '100%',
  borderRadius: 4,
  padding: '0 10px',
  fontSize: 15,
  lineHeight: 1,
  color: '#000',
  // boxShadow: `0 0 0 1px #000`,
  border: '1px solid #949494',
  height: 40,
  marginBottom: '8px',
  '&:focus': { boxShadow: `0 0 0 1px #646464` },
})

export const Brand = styled('div', {
  width: '120px',
  margin: '0 auto',
  marginBottom: '22px',
  img: {
    width: '100%',
  },
})
