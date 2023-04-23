import useSWR from 'swr'

const fetcher = (...args) => fetch(...args).then((res) => res.json())
// const headers = {
//   headers: {
//     apikey: process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY,
//     Authorization: 'Bearer ' + process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY,
//   },
// }

function useRequest(url) {
  const { data, error } = useSWR(process.env.NEXT_PUBLIC_STRAPI_API_URL + url, fetcher)
  return {
    data: data,
    loading: !error && !data,
    error: error,
  }
}

function useRequestConditional(url, filter) {
  const { data, error } = useSWR(
    filter ? process.env.NEXT_PUBLIC_STRAPI_API_URL + url : null,
    fetcher
  )
  return {
    data: data,
    loading: !error && !data,
    error: error,
  }
}

export { useRequest, useRequestConditional }
