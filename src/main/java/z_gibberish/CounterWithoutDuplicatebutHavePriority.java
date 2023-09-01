package z_gibberish;
import java.util.*;

public class CounterWithoutDuplicatebutHavePriority {

    public static void main(String[] args) {
        // Create TreeSet for numb and result
        TreeMap<Integer, String> numbTreeSet = new TreeMap<> ();

        // Sample data
        int[] numbs = {1, 1, 1, 2};
        String[] results = {"A", "B", "C", "C"};

        // Iterate through the data
        for (int i = 0; i < numbs.length; i++) {
            int numb = numbs[i];
            String result = results[i];

            // Check if numb already exists in the TreeSet
            if (numbTreeSet.containsKey (numb)) {
                // Check if the result is "C", if so, update it
                if ("C".equalsIgnoreCase (result)) {
                    numbTreeSet.put (numb, "C");
                }
            } else {
                // If numb doesn't exist, add it to the TreeSet
                numbTreeSet.put (numb, result);
            }
        }

        // Print the TreeSet contents
        for (Map.Entry<Integer, String> entry : numbTreeSet.entrySet ()) {
            int numb = entry.getKey ();
            String result = entry.getValue ();
            System.out.println ("numb: " + numb + ", result: " + result);
        }
    }
}
